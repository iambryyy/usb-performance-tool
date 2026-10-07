import atexit
import csv
import ctypes
import functools
import hashlib
import json
import os
import platform
import re
import shutil
import socket
import statistics
import subprocess
import sys
import tempfile
import threading
import time
import logging
import webbrowser
from concurrent.futures import ThreadPoolExecutor
from dataclasses import asdict, dataclass
from datetime import datetime
from html import escape as esc
from pathlib import Path
from urllib.error import URLError
from urllib.request import urlopen

from app_paths import app_dir, resolve_path

if sys.platform == "win32":
	# The packaged program is a windowed app with no console. Every console helper it
	# launches (powershell, ping, netsh, net view...) therefore gets a console of its own,
	# which Windows opens as a visible Windows Terminal window that steals keyboard focus.
	# The test makes dozens of these calls before its first progress bar, so the screen
	# flickered and typing was interrupted. Start every child process without a window.
	_popen_init = subprocess.Popen.__init__

	def _windowless_popen_init(self, *args, **kwargs):
		kwargs["creationflags"] = kwargs.get("creationflags", 0) | subprocess.CREATE_NO_WINDOW
		_popen_init(self, *args, **kwargs)

	subprocess.Popen.__init__ = _windowless_popen_init



MB = 1024 * 1024


@dataclass
class Result:
	project_name: str
	dut_model: str
	dut_role: str
	firmware_version: str
	topology: str
	client_type: str
	client_os: str
	client_ip: str
	configured_interface: str
	connection_type: str
	actual_ssid: str
	actual_channel: str
	actual_band: str
	detected_gateway_ip: str = ""
	dut_name: str = ""
	detected_hardware_version: str = ""
	bssid: str = ""
	wifi_signal_percent: str = ""
	usb_share_path: str = ""
	actual_usb_format: str = ""
	expected_interface: str = ""
	interface_match: bool = False
	role_assumed: bool = False
	dut_reachable: bool = False
	share_write_access: bool = False
	usb_format: str = ""
	direction: str = ""
	file_set_version: str = ""
	file_name: str = ""
	iteration: int = 0
	throughput_mbps: float = 0.0
	duration_seconds: float = 0.0
	average_mbps: float = 0.0
	min_mbps: float = 0.0
	max_mbps: float = 0.0
	stddev_mbps: float = 0.0
	passed: bool = False
	status: str = "NOT_RUN"
	failure_reason: str = ""
	timestamp: str = ""
	file_verification: str = ""


def load_config(path="test_config.json"):
	path = resolve_config_path(path)
	with open(path, "r", encoding="utf-8") as config_file:
		return json.load(config_file)


def resolve_config_path(path="test_config.json"):
	config_path = Path(path)
	candidates = [config_path]
	if not config_path.is_absolute():
		candidates.append(Path(app_dir()) / config_path)
	for candidate in list(candidates):
		candidates.append(candidate.with_name(candidate.stem + ".example" + candidate.suffix))
	for candidate in candidates:
		if candidate.is_file():
			return candidate
	return config_path


def is_unknown(value):
	return str(value or "").strip().lower() in ["", "auto", "unknown", "n/a", "none"]


def _can_prompt():
	try:
		return sys.stdin is not None and sys.stdin.isatty()
	except (OSError, ValueError):
		return False


def prompt_value(label, default=""):
	default = "" if is_unknown(default) else str(default)
	# When run from the GUI there is no console to type into; input() would
	# fail and end the whole run. Use the default and continue instead.
	if not _can_prompt():
		print(f"{label}: no console to ask, using {default or '(nothing)'}")
		return default
	prompt = label
	if default:
		prompt += f" [{default}]"
	try:
		answer = input(prompt + ": ").strip()
	except EOFError:
		answer = ""
	return answer or default


def prompt_choice(label, choices, default=""):
	choices_text = "/".join(choices)
	while True:
		answer = prompt_value(f"{label} ({choices_text})", default).strip()
		if not answer:
			return ""
		for choice in choices:
			if answer.lower() == choice.lower():
				return choice
		print(f"Please enter one of: {choices_text}")


def normalize_share_path(value, share_name="g"):
	value = value.strip()
	if not value:
		return ""
	if value.startswith("\\\\"):
		return value
	return f"\\\\{value}\\{share_name}"


def ensure_dir(path):
	Path(path).mkdir(parents=True, exist_ok=True)

LOG_PATH = ""


def setup_logging():
	global LOG_PATH
	log_dir = resolve_path("logs")
	ensure_dir(log_dir)
	LOG_PATH = os.path.join(log_dir, f"usb_test_{datetime.now():%Y%m%d_%H%M%S}.log")
	logging.basicConfig(
		filename=LOG_PATH,
		level=logging.INFO,
		format="%(asctime)s %(levelname)s %(message)s",
	)

_ip_probe_target = "8.8.8.8"


def client_ip():
	# UDP connect() sends nothing; it just makes the OS pick the local address
	# it would use to reach the target. Probing toward the DUT (set by run())
	# rather than the internet gives the client's address on the network
	# actually under test, even when another connection (Ethernet, VPN) has the
	# preferred internet route.
	try:
		with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
			sock.connect((_ip_probe_target, 80))
			return sock.getsockname()[0]
	except OSError:
		try:
			return socket.gethostbyname(socket.gethostname())
		except OSError:
			return ""


def client_ip_problem(ip):
	"""Why this client address can't be used for a test, or "" if it is fine.
	169.254.x.x is what Windows assigns itself when DHCP gave it nothing."""
	if not ip:
		return "this PC has no IP address on the network under test"
	if ip.startswith("169.254."):
		return f"this PC only has a self-assigned address ({ip}), so DHCP gave it no IP"
	if ip.startswith("127."):
		return f"this PC has no network address (only {ip})"
	return ""


def detect_client_type():
	"""Auto-detect client OS type from platform."""
	return {"Windows": "Windows PC", "Darwin": "Mac", "Linux": "Linux PC"}.get(
		platform.system(), platform.system() or "Unknown"
	)


def detect_client_os():
	"""
	Auto-detect a clean, human-readable OS label (e.g. "Windows 11"),
	without kernel/build noise (e.g. "Windows-10-10.0.22631-SP0").
	Windows 11 reports the same major/minor version as Windows 10 in
	platform.version() -- only the build number tells them apart, so a
	raw platform.platform() string can't distinguish the two reliably.
	"""
	system = platform.system()
	if system == "Windows":
		try:
			build = int(platform.version().split(".")[-1])
		except (ValueError, IndexError):
			build = 0
		# Builds 22000+ are Windows 11; below that, Windows 10.
		version = "11" if build >= 22000 else "10"
		return f"Windows {version}"
	if system == "Darwin":
		return f"macOS {platform.mac_ver()[0]}".strip()
	if system == "Linux":
		return f"Linux {platform.release().split('-')[0]}"
	return system or "Unknown"


def gateway_ip(near=""):
	"""
	Default gateway of the network under test. With several connections up at
	once (Ethernet/VPN plus WiFi), the lowest-metric default route usually
	belongs to the wrong one, so when `near` (the USB share's host) is given,
	prefer the gateway on the same /24 as it.
	"""
	try:
		result = subprocess.run(
			[
				"powershell",
				"-NoProfile",
				"-NonInteractive",
				"-Command",
				"Get-NetRoute -DestinationPrefix '0.0.0.0/0' | "
				"Sort-Object RouteMetric | "
				"ForEach-Object { $_.NextHop }",
			],
			capture_output=True,
			text=True,
			check=False,
		)
		gateways = [g.strip() for g in result.stdout.splitlines() if g.strip()]
		if near and re.match(r"^\d+\.\d+\.\d+\.\d+$", near):
			prefix = near.rsplit(".", 1)[0] + "."
			for gateway in gateways:
				if gateway.startswith(prefix):
					return gateway
		if gateways:
			return gateways[0]
	except Exception:
		pass
	return ""


def http_text(url, timeout=3):
	try:
		with urlopen(url, timeout=timeout) as response:
			return response.read().decode("utf-8", errors="ignore")
	except (OSError, URLError):
		return ""


def resolve_host_ip(host):
	"""Resolve a hostname to its IP. Returns host unchanged if already an IP or lookup fails."""
	if not host:
		return ""
	if re.match(r"\d{1,3}(?:\.\d{1,3}){3}$", host):
		return host
	try:
		return socket.gethostbyname(host)
	except (socket.gaierror, OSError):
		return host



def infer_dut_role(dut_ip_address, gateway):
	"""'Router' if the USB share's host is the network's default gateway, otherwise 'Extender'."""
	if not dut_ip_address or not gateway:
		return "Unknown"
	return "Router" if resolve_host_ip(dut_ip_address) == gateway else "Extender"


def resolve_dut_role(config, ip):
	"""
	Determine the device role and whether it was confirmed or assumed. A role set in
	the configuration wins; otherwise the share's host is compared with the default
	gateway (the main router is the gateway; any other device is an extender).
	"""
	configured_role = config.get("dut", {}).get("role", "auto")
	if not is_unknown(configured_role):
		return configured_role, False
	share_ip = resolve_host_ip(share_host(config.get("usb_share_path", "")))
	return infer_dut_role(share_ip or ip, config.get("detected_gateway_ip", "")), False


def upnp_device_info(host, timeout=3,
					 candidates=((49000, "rootDesc.xml"), (5000, "rootDesc.xml"), (80, "rootDesc.xml"), (49152, "gatedesc.xml"))):
	"""
	Read what a router volunteers about itself over UPnP (its device-description XML):
	friendly name, manufacturer, model and version. Returns {} if it does not answer.
	Only ports that actually accept a connection are tried, so a quiet device costs
	well under a second.
	"""
	import xml.etree.ElementTree as ET
	for port, path in candidates:
		try:
			socket.create_connection((host, port), timeout=0.5).close()
		except OSError:
			continue
		text = http_text(f"http://{host}:{port}/{path}", timeout=timeout)
		if "<device" not in text:
			continue
		try:
			device = ET.fromstring(text).find(".//{*}device")
		except ET.ParseError:
			continue
		if device is None:
			continue

		def field(name):
			element = device.find("{*}" + name)
			return (element.text or "").strip() if element is not None and element.text else ""

		maker, model, number = field("manufacturer"), field("modelName"), field("modelNumber")
		info = {"name": field("friendlyName"),
				"model": f"{maker} {model}".strip() if maker and maker.lower() not in model.lower() else model}
		if re.match(r"^\d+(\.\d+)+", number):
			info["firmware_version"] = number
		return {k: v for k, v in info.items() if v}
	return {}


def detect_dut_info(config):
	"""
	The device under test: role, name, model and firmware. Values in the configuration
	win; anything left unknown is read from the device itself over UPnP when it
	offers that; whatever still cannot be read is reported as Unknown, never guessed.
	"""
	ip = config.get("detected_gateway_ip", "") or dut_ip(config)
	role, role_assumed = resolve_dut_role(config, ip)
	dut = dict(config.get("dut", {}))
	dut["role"] = role
	dut["role_assumed"] = role_assumed
	dut.setdefault("info_source", "config")
	configured_ip = dut.get("ip_address", "")
	host = "" if is_unknown(configured_ip) else configured_ip
	host = host or share_host(config.get("usb_share_path", "")) or config.get("detected_gateway_ip", "")
	wanted = [key for key in ("name", "model", "firmware_version") if is_unknown(dut.get(key, ""))]
	# Only probe a literal IP address: a host name needs a DNS/NetBIOS lookup that
	# the HTTP timeout does not cover and that can stall for a long time.
	if wanted and re.match(r"^\d+\.\d+\.\d+\.\d+$", host or ""):
		info = upnp_device_info(host)
		for key in wanted:
			if info.get(key):
				dut[key] = info[key]
				dut["info_source"] = "upnp"
	for key in ("model", "firmware_version"):
		if is_unknown(dut.get(key, "")):
			dut[key] = "Unknown"
	if is_unknown(dut.get("name", "")):
		dut["name"] = ""
	return dut


def share_host(share):
	match = re.search(r"\\\\([^\\]+)\\", share or "")
	return match.group(1) if match else ""


def target_share_path(config, target):
	share = target.get("usb_share_path", "")
	if share and share.lower() != "auto":
		return share
	ip = target.get("ip_address", "")
	share_name = target.get("usb_share_name") or config.get("usb_share_name", "g")
	if ip and ip.lower() != "auto":
		return f"\\\\{ip}\\{share_name}"
	return ""


def target_candidates(config):
	targets = config.get("dut_targets") or []
	if targets:
		return [dict(target) for target in targets]
	dut = dict(config.get("dut", {}))
	share = config.get("usb_share_path", "")
	dut["usb_share_path"] = share
	dut.setdefault("ip_address", share_host(share))
	return [dut]


def find_target(config, target_name):
	for target in target_candidates(config):
		if target_matches(target, target_name.lower()):
			return target
	return {}


def target_matches(target, requested):
	values = [target.get("name", ""), target.get("role", ""), target.get("ip_address", "")]
	return any(str(value).lower() == requested for value in values)


def apply_target(config, target, source):
	updated = dict(config)
	dut = dict(updated.get("dut", {}))
	share = target_share_path(updated, target)
	for key in ["model", "role", "firmware_version", "ip_address"]:
		if target.get(key) and str(target.get(key)).lower() != "auto":
			dut[key] = target[key]
	if not dut.get("ip_address") and share:
		dut["ip_address"] = share_host(share)
	dut["info_source"] = source
	updated["dut"] = dut
	if share:
		updated["usb_share_path"] = share
	updated["selected_target"] = target.get("name") or target.get("role") or dut.get("ip_address", "Unknown")
	return updated


def resolve_target_dut(config):
	requested = str(config.get("target_dut", "auto")).lower()
	candidates = target_candidates(config)
	if requested != "auto":
		matches = [target for target in candidates if target_matches(target, requested)]
		if matches:
			return apply_target(config, matches[0], "target-config")
		config = dict(config)
		config["target_resolution_error"] = f"requested target_dut '{requested}' was not found in dut_targets"
		return config
	for target in candidates:
		share = target_share_path(config, target)
		if share and validate_share_write_access(share):
			return apply_target(config, target, "auto-selected-writable-share")
	config = dict(config)
	dut = dict(config.get("dut", {}))
	dut["role"] = "Unknown"
	dut["info_source"] = "not-selected-no-writable-usb-share"
	config["dut"] = dut
	config["target_resolution_error"] = "no configured DUT target has a writable USB share"
	return config


def prompt_for_run_context(config):
	if not config.get("prompt_for_run_context", False):
		return config
	config = dict(config)
	target = prompt_choice("DUT target for this run", ["Router", "Extender", "auto"], config.get("target_dut", "auto"))
	if target:
		config["target_dut"] = target
	if target and target.lower() != "auto":
		selected = find_target(config, target)
		share = target_share_path(config, selected) if selected else ""
		if is_unknown(share) or not os.path.isdir(share):
			share = normalize_share_path(prompt_value(f"{target} USB share path or DUT IP", share), config.get("usb_share_name", "g"))
			if share:
				config["usb_share_path"] = share
				if selected:
					selected["usb_share_path"] = share
					selected["ip_address"] = share_host(share)
					config["dut_targets"] = [selected if target_matches(item, target.lower()) else item for item in target_candidates(config)]
	return config


def derived_topology(config):
	dut = config.get("dut", {})
	role = dut.get("role") or config.get("selected_target") or "Unknown DUT"
	connection = config.get("actual_connection_type", "unknown")
	wifi = config.get("actual_wifi", {})
	# The client is often attached to a different device than the one holding the
	# drive (drive in the Router, cable in an Extender); say so only when it was declared.
	attached = str(config.get("client_attached_to", "")).strip().upper()
	where = f" connected to {attached}" if attached in ("Router", "Extender") and attached != str(role).upper() else ""
	if connection == "wireless" and wifi.get("band"):
		return f"{role} USB performance test - wireless client{where} on {wifi.get('band')}"
	if connection == "wired":
		return f"{role} USB performance test - wired client{where}"
	return f"{role} USB performance test"


def prompt_for_missing_info(config):
	if not config.get("prompt_for_missing_info", True):
		return config
	config = dict(config)
	dut = dict(config.get("dut", {}))
	changed = False
	if is_unknown(dut.get("role")):
		role = prompt_value("DUT role (Router or Extender)", config.get("selected_target", ""))
		if role:
			dut["role"] = role.upper()
			changed = True
	if is_unknown(dut.get("model")):
		model = prompt_value("DUT model", dut.get("model", ""))
		if model:
			dut["model"] = model
			changed = True
	if is_unknown(dut.get("firmware_version")):
		firmware = prompt_value("DUT firmware version", dut.get("firmware_version", ""))
		if firmware:
			dut["firmware_version"] = firmware
			changed = True
	share = config.get("usb_share_path", "")
	if is_unknown(share) or not os.path.isdir(share):
		share = normalize_share_path(prompt_value("USB share path or DUT IP", share), config.get("usb_share_name", "g"))
		if share:
			config["usb_share_path"] = share
			dut["ip_address"] = share_host(share)
			changed = True
	if changed:
		dut["info_source"] = "user-provided"
	config["dut"] = dut
	return config


def load_file_manifest():
	path = Path(app_dir()) / "file_set_manifest.json"
	try:
		return json.loads(path.read_text(encoding="utf-8"))
	except (OSError, ValueError):
		return None


def sha256_of(path):
	digest = hashlib.sha256()
	with open(path, "rb") as f:
		for block in iter(lambda: f.read(4 * MB), b""):
			digest.update(block)
	return digest.hexdigest()


def verify_test_files(config, files):
	"""
	Tag each test file as verified / CHECKSUM MISMATCH / custom against the
	standard file set manifest, and return a one-line summary for the report.
	"""
	manifest = load_file_manifest()
	if not manifest:
		for item in files:
			item["verification"] = "unverified (no manifest)"
		return "no file_set_manifest.json found; file set not verified"
	folder = resolve_path(config.get("local_test_file_dir", "test_files"))
	expected = {entry["name"]: entry for entry in manifest.get("files", [])}
	verified, mismatched, custom, unverified = [], [], [], []
	for item in files:
		entry = expected.get(item["name"])
		path = os.path.join(folder, item["name"])
		if not entry:
			item["verification"] = "custom (not in standard set)"
			custom.append(item["name"])
		elif not os.path.exists(path):
			item["verification"] = "unverified (not present locally)"
			unverified.append(item["name"])
		elif os.path.getsize(path) != int(entry["size_mb"]) * MB or sha256_of(path) != entry["sha256"]:
			item["verification"] = "CHECKSUM MISMATCH"
			mismatched.append(item["name"])
		else:
			item["verification"] = "verified"
			verified.append(item["name"])
	present = {item["name"] for item in files}
	missing = [name for name in expected if name not in present]
	parts = [f"Standard set {manifest.get('version', '?')}: {len(verified)} of {len(files)} file(s) used in this run verified"]
	if mismatched:
		parts.append("MISMATCH: " + ", ".join(mismatched) + " (run generate_test_files.py)")
	if missing:
		parts.append("standard files not used in this run: " + ", ".join(missing))
	if custom:
		parts.append("custom files: " + ", ".join(custom))
	if unverified:
		parts.append("unverified: " + ", ".join(unverified))
	if config.get("test_file_set_version") != manifest.get("version"):
		parts.append(f"config test_file_set_version '{config.get('test_file_set_version')}' differs from manifest '{manifest.get('version')}'")
	return "; ".join(parts)


def discover_test_files(config):
	configured = config.get("test_files", "auto")
	if configured != "auto":
		return configured
	folder = resolve_path(config.get("local_test_file_dir", "test_files"))
	ensure_dir(folder)
	files = [
		{"name": p.name, "size_mb": round(p.stat().st_size / MB, 2), "type": "discovered_file", "generate": False}
		for p in sorted(Path(folder).iterdir()) if p.is_file() and p.stat().st_size > 0
	]
	if not files:
		raise FileNotFoundError(f"No test files found in {folder}. Add one or set test_files explicitly.")
	return files


def powershell(command, timeout=5):
	try:
		return subprocess.run(
			["powershell", "-NoProfile", "-Command", command],
			capture_output=True,
			text=True,
			timeout=timeout,
			check=False,
		).stdout
	except (OSError, subprocess.TimeoutExpired):
		return ""


def wifi_info():
    if platform.system() != "Windows":
        return {
            "ssid": "",
            "bssid": "",
            "signal": "",
            "channel": "",
            "band": "",
            "connected": False
        }

    try:
        output = subprocess.run(
            ["netsh", "wlan", "show", "interfaces"],
            capture_output=True,
            text=True,
            timeout=5,
            check=False,
        ).stdout
    except (OSError, subprocess.TimeoutExpired):
        output = ""

    lower = output.lower()

    connected = bool(
        re.search(
            r"^\s*state\s*:\s*connected\s*$",
            lower,
            re.MULTILINE
        )
    )

    data = {
        "ssid": "",
        "bssid": "",
        "signal": "",
        "channel": "",
        "band": "",
        "connected": connected,
    }

    for key, pattern in [
        ("ssid",    r"^\s*SSID\s*:\s*(.+?)\s*$"),
        ("channel", r"^\s*Channel\s*:\s*(.+?)\s*$"),
        ("band",    r"^\s*Band\s*:\s*(.+?)\s*$"),
        ("bssid",   r"^\s*AP BSSID\s*:\s*([0-9a-f:]+)"),
        ("signal",  r"^\s*Signal\s*:\s*(.+?)\s*$"),
    ]:
        m = re.search(pattern, output, re.MULTILINE | re.IGNORECASE)
        if m:
            data[key] = m.group(1).strip()

    if not data["band"] and data["channel"]:
        try:
            ch = int(data["channel"])

            if ch <= 14:
                data["band"] = "2.4GHz"
            elif ch >= 180:
                data["band"] = "6GHz"
            else:
                data["band"] = "5GHz"

        except ValueError:
            pass

    return data


def create_temp_wifi_profile(ssid, password, security="WPA2PSK"):
	"""
	Install a temporary Windows WLAN profile for a network the laptop
	doesn't already have saved, using a typed SSID/password.

	The generated profile XML is written to a short-lived temp file only
	long enough for `netsh wlan add profile` to read it, then deleted
	immediately -- the password is never written to any file this program
	controls (test_config.json, logs, etc.). Windows itself will still
	store the profile (and its credentials) in its own profile store once
	installed; call remove_temp_wifi_profile() after use to clean that up
	too.

	Assumes WPA2-Personal (the common case for typical test networks).
	Returns False for open/WPA3/enterprise networks this doesn't support.
	"""
	if platform.system() != "Windows":
		return False
	if not ssid or not password:
		return False

	import xml.sax.saxutils as xml_escape
	ssid_safe = xml_escape.escape(ssid)
	password_safe = xml_escape.escape(password)
	profile_xml = f"""<?xml version="1.0"?>
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
	<name>{ssid_safe}</name>
	<SSIDConfig>
		<SSID>
			<name>{ssid_safe}</name>
		</SSID>
	</SSIDConfig>
	<connectionType>ESS</connectionType>
	<connectionMode>manual</connectionMode>
	<MSM>
		<security>
			<authEncryption>
				<authentication>{security}</authentication>
				<encryption>AES</encryption>
				<useOneX>false</useOneX>
			</authEncryption>
			<sharedKey>
				<keyType>passPhrase</keyType>
				<protected>false</protected>
				<keyMaterial>{password_safe}</keyMaterial>
			</sharedKey>
		</security>
	</MSM>
</WLANProfile>"""

	no_window = subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0
	tmp_path = None
	try:
		with tempfile.NamedTemporaryFile(mode="w", suffix=".xml", delete=False, encoding="utf-8") as f:
			f.write(profile_xml)
			tmp_path = f.name
		result = subprocess.run(
			["netsh", "wlan", "add", "profile", f"filename={tmp_path}", "user=all"],
			capture_output=True, text=True, timeout=10, check=False,
			creationflags=no_window,
		)
		return result.returncode == 0
	except (OSError, subprocess.TimeoutExpired):
		return False
	finally:
		if tmp_path and os.path.exists(tmp_path):
			try:
				os.remove(tmp_path)
			except OSError:
				pass


def remove_temp_wifi_profile(ssid):
	"""
	Remove a temporary WiFi profile after use, so its password doesn't
	linger in Windows' own saved-profile store beyond this test run.
	"""
	if platform.system() != "Windows" or not ssid:
		return
	no_window = subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0
	try:
		subprocess.run(
			["netsh", "wlan", "delete", "profile", f"name={ssid}"],
			capture_output=True, text=True, timeout=10, check=False,
			creationflags=no_window,
		)
	except (OSError, subprocess.TimeoutExpired):
		pass


def parse_ssid_bssids(text, ssid):
	"""Parse `netsh wlan show networks mode=bssid` into [{bssid, band, signal}] for one SSID."""
	found = []
	in_ssid = False
	current = None
	for line in text.splitlines():
		m = re.match(r"^SSID\s+\d+\s*:\s*(.*?)\s*$", line)
		if m:
			in_ssid = m.group(1) == ssid
			current = None
			continue
		if not in_ssid:
			continue
		m = re.match(r"^\s*BSSID\s+\d+\s*:\s*([0-9a-fA-F:]{17})", line)
		if m:
			current = {"bssid": m.group(1).lower(), "band": "", "signal": 0}
			found.append(current)
			continue
		if current is None:
			continue
		m = re.match(r"^\s*Signal\s*:\s*(\d+)%", line)
		if m:
			current["signal"] = int(m.group(1))
			continue
		m = re.match(r"^\s*Band\s*:\s*(.+?)\s*$", line)
		if m:
			current["band"] = m.group(1).replace(" ", "").lower()
	return found


@functools.lru_cache(maxsize=None)
def _wlan_api():
	"""ctypes definitions for the few Native Wifi calls needed to pin a connection to one BSSID."""
	from ctypes import wintypes
	api = ctypes.WinDLL("wlanapi")

	class GUID(ctypes.Structure):
		_fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD), ("Data3", wintypes.WORD), ("Data4", ctypes.c_ubyte * 8)]

	class INTERFACE_INFO(ctypes.Structure):
		_fields_ = [("InterfaceGuid", GUID), ("strInterfaceDescription", ctypes.c_wchar * 256), ("isState", wintypes.DWORD)]

	class INTERFACE_LIST(ctypes.Structure):
		_fields_ = [("dwNumberOfItems", wintypes.DWORD), ("dwIndex", wintypes.DWORD), ("InterfaceInfo", INTERFACE_INFO * 1)]

	class DOT11_SSID(ctypes.Structure):
		_fields_ = [("uSSIDLength", wintypes.ULONG), ("ucSSID", ctypes.c_ubyte * 32)]

	class NDIS_OBJECT_HEADER(ctypes.Structure):
		_fields_ = [("Type", ctypes.c_ubyte), ("Revision", ctypes.c_ubyte), ("Size", ctypes.c_ushort)]

	class DOT11_BSSID_LIST(ctypes.Structure):
		_fields_ = [("Header", NDIS_OBJECT_HEADER), ("uNumOfEntries", wintypes.ULONG),
		            ("uTotalNumOfEntries", wintypes.ULONG), ("BSSIDs", (ctypes.c_ubyte * 6) * 1)]

	class CONNECTION_PARAMETERS(ctypes.Structure):
		_fields_ = [("wlanConnectionMode", wintypes.DWORD), ("strProfile", ctypes.c_wchar_p),
		            ("pDot11Ssid", ctypes.POINTER(DOT11_SSID)), ("pDesiredBssidList", ctypes.POINTER(DOT11_BSSID_LIST)),
		            ("dot11BssType", wintypes.DWORD), ("dwFlags", wintypes.DWORD)]

	api.WlanOpenHandle.argtypes = [wintypes.DWORD, ctypes.c_void_p, ctypes.POINTER(wintypes.DWORD), ctypes.POINTER(wintypes.HANDLE)]
	api.WlanOpenHandle.restype = wintypes.DWORD
	api.WlanCloseHandle.argtypes = [wintypes.HANDLE, ctypes.c_void_p]
	api.WlanCloseHandle.restype = wintypes.DWORD
	api.WlanEnumInterfaces.argtypes = [wintypes.HANDLE, ctypes.c_void_p, ctypes.POINTER(ctypes.POINTER(INTERFACE_LIST))]
	api.WlanEnumInterfaces.restype = wintypes.DWORD
	api.WlanFreeMemory.argtypes = [ctypes.c_void_p]
	api.WlanScan.argtypes = [wintypes.HANDLE, ctypes.POINTER(GUID), ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p]
	api.WlanScan.restype = wintypes.DWORD
	api.WlanConnect.argtypes = [wintypes.HANDLE, ctypes.POINTER(GUID), ctypes.POINTER(CONNECTION_PARAMETERS), ctypes.c_void_p]
	api.WlanConnect.restype = wintypes.DWORD
	return api, GUID, INTERFACE_LIST, DOT11_SSID, NDIS_OBJECT_HEADER, DOT11_BSSID_LIST, CONNECTION_PARAMETERS


def _wlan_session(action):
	"""Open a WLAN handle on the first wireless interface, run action(api, handle, guid_ptr), always clean up."""
	if platform.system() != "Windows":
		return None
	api, GUID, INTERFACE_LIST, *_ = _wlan_api()
	from ctypes import wintypes
	negotiated = wintypes.DWORD()
	handle = wintypes.HANDLE()
	if api.WlanOpenHandle(2, None, ctypes.byref(negotiated), ctypes.byref(handle)) != 0:
		return None
	iface_list = ctypes.POINTER(INTERFACE_LIST)()
	try:
		if api.WlanEnumInterfaces(handle, None, ctypes.byref(iface_list)) != 0 or iface_list.contents.dwNumberOfItems < 1:
			return None
		guid = GUID.from_buffer_copy(iface_list.contents.InterfaceInfo[0].InterfaceGuid)
		return action(api, handle, ctypes.byref(guid))
	finally:
		if iface_list:
			api.WlanFreeMemory(iface_list)
		api.WlanCloseHandle(handle, None)


def _wlan_connect_bssid(ssid, bssid):
	"""Ask Windows to connect the saved profile `ssid` to one specific access point radio (no Administrator needed)."""
	api, GUID, INTERFACE_LIST, DOT11_SSID, NDIS_OBJECT_HEADER, DOT11_BSSID_LIST, CONNECTION_PARAMETERS = _wlan_api()
	raw = bytes(int(part, 16) for part in bssid.split(":"))
	ssid_struct = DOT11_SSID()
	encoded = ssid.encode("utf-8")[:32]
	ssid_struct.uSSIDLength = len(encoded)
	for i, byte in enumerate(encoded):
		ssid_struct.ucSSID[i] = byte
	bssid_list = DOT11_BSSID_LIST()
	bssid_list.Header.Type = 0x80          # NDIS_OBJECT_TYPE_DEFAULT
	bssid_list.Header.Revision = 1         # DOT11_BSSID_LIST_REVISION_1
	bssid_list.Header.Size = ctypes.sizeof(DOT11_BSSID_LIST)
	bssid_list.uNumOfEntries = 1
	bssid_list.uTotalNumOfEntries = 1
	for i, byte in enumerate(raw):
		bssid_list.BSSIDs[0][i] = byte
	params = CONNECTION_PARAMETERS()
	params.wlanConnectionMode = 0          # wlan_connection_mode_profile
	params.strProfile = ssid
	params.pDot11Ssid = ctypes.pointer(ssid_struct)
	params.pDesiredBssidList = ctypes.pointer(bssid_list)
	params.dot11BssType = 1                # dot11_BSS_type_infrastructure
	params.dwFlags = 0
	result = _wlan_session(lambda a, h, g: a.WlanConnect(h, g, ctypes.byref(params), None))
	return result == 0


def _wlan_scan():
	_wlan_session(lambda a, h, g: a.WlanScan(h, g, None, None, None))


_switch_detail = ""


def connect_pinned_to_band(ssid, expected_band, wait=20):
	"""
	Reach a specific band of a shared SSID by connecting to that band's radio
	(BSSID) directly. Works without Administrator rights and on any adapter.
	"""
	global _switch_detail
	want = expected_band.replace(" ", "").lower()
	seen = []
	# A Wi-Fi scan takes several seconds (longer with a 6GHz radio), and a
	# radio can be missing from the first result, so scan up to 3 times
	# before concluding the band isn't there.
	for _ in range(3):
		try:
			_wlan_scan()
			time.sleep(4)
			scan = subprocess.run(["netsh", "wlan", "show", "networks", "mode=bssid"], capture_output=True, text=True, timeout=15, check=False).stdout
		except (OSError, subprocess.TimeoutExpired):
			_switch_detail = "the Wi-Fi scan failed"
			return False
		seen = parse_ssid_bssids(scan, ssid)
		if any(b["band"] == want for b in seen):
			break
	candidates = sorted((b for b in seen if b["band"] == want), key=lambda b: -b["signal"])
	if not candidates:
		bands = sorted({b["band"] for b in seen})
		_switch_detail = (
			f"no {expected_band} radio for '{ssid}' was visible from here; radios seen: "
			+ (", ".join(bands) if bands else "none")
		)
		return False
	for candidate in candidates[:3]:
		try:
			if not _wlan_connect_bssid(ssid, candidate["bssid"]):
				continue
		except (OSError, ValueError, AttributeError):
			_switch_detail = "Windows could not start the connection"
			return False
		waited = 0
		while waited < wait:
			time.sleep(1)
			waited += 1
			info = wifi_info()
			if info.get("connected") and info.get("ssid") == ssid and info.get("band", "").replace(" ", "").lower() == want:
				return True
	_switch_detail = f"a {expected_band} radio was visible ({len(candidates)} found) but Windows did not join it"
	return False


_BAND_PREF_PATTERNS = {
	"2.4ghz": r"^\d+\.\s*Prefer 2\.4\s?GHz",
	"5ghz": r"^\d+\.\s*Prefer 5\s?GHz band$",
	"6ghz": r"^\d+\.\s*Prefer 6\s?GHz",
}
_original_band_pref = {}


def _wifi_adapter_name():
	try:
		out = subprocess.run(["netsh", "wlan", "show", "interfaces"], capture_output=True, text=True, timeout=5, check=False).stdout
	except (OSError, subprocess.TimeoutExpired):
		return ""
	m = re.search(r"^\s*Name\s*:\s*(.+?)\s*$", out, re.MULTILINE)
	return m.group(1) if m else ""


def _set_band_preference(display_value=None, band=None):
	"""
	Set the WiFi adapter's "Preferred Band" driver property (Intel drivers).
	Pass band ('2.4GHz'/'5GHz'/'6GHz') to pick the matching value, or
	display_value to restore an exact one. Needs an Administrator session.
	Returns the previous display value on success, or None on failure.
	"""
	name = _wifi_adapter_name()
	if platform.system() != "Windows" or not name:
		return None
	pattern = _BAND_PREF_PATTERNS.get((band or "").replace(" ", "").lower(), "")
	script = (
		f"$n='{name}'; $p=Get-NetAdapterAdvancedProperty -Name $n -DisplayName 'Preferred Band' -ErrorAction Stop; "
		f"$want='{display_value or ''}'; "
		f"if (-not $want) {{ $want=($p.ValidDisplayValues | Where-Object {{ $_ -match '{pattern}' }} | Select-Object -First 1) }}; "
		"if (-not $want) { exit 3 }; "
		"Set-NetAdapterAdvancedProperty -Name $n -DisplayName 'Preferred Band' -DisplayValue $want -ErrorAction Stop; "
		"Write-Output $p.DisplayValue"
	)
	try:
		result = subprocess.run(
			["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
			capture_output=True, text=True, timeout=45, check=False,
		)
	except (OSError, subprocess.TimeoutExpired):
		return None
	if result.returncode != 0:
		return None
	return result.stdout.strip() or None


def restore_band_preference():
	previous = _original_band_pref.pop("value", None)
	if previous:
		_set_band_preference(display_value=previous)


def switch_wifi_band(ssid, expected_band, timeout=25, poll_interval=1, password=None):
	"""
	Switch the client's WiFi connection to a specific band's SSID, and
	verify the switch actually landed on the expected band before
	returning -- not just that *some* connection succeeded.

	If password is given and no profile for this SSID already exists,
	a temporary profile is created first (see create_temp_wifi_profile).
	If password is omitted, this requires a profile to already exist on
	this laptop (i.e. it has connected to that SSID before) -- matching
	the original behavior, unchanged for anyone not using the password
	path.

	On routers with band-steering enabled, connecting to a shared SSID
	doesn't guarantee which band you land on -- the router decides, not
	the client. If the actual connected band doesn't match expected_band,
	this returns False rather than silently mislabeling results as the
	wrong band.

	MLO band reporting via netsh is newer and less consistent across
	Windows/driver versions than plain 2.4/5/6GHz, so for MLO this only
	verifies the connection and SSID, not an exact band string match.
	"""
	global _switch_detail
	_switch_detail = ""
	if platform.system() != "Windows":
		return False
	if not ssid:
		return False
	if password:
		if not create_temp_wifi_profile(ssid, password):
			_switch_detail = "the Wi-Fi profile could not be created (check the password)"
			return False
	def connect_and_check(wait):
		try:
			subprocess.run(
				["netsh", "wlan", "connect", f"name={ssid}", f"ssid={ssid}"],
				capture_output=True, text=True, timeout=10, check=False,
			)
		except (OSError, subprocess.TimeoutExpired):
			return False
		waited = 0
		while waited < wait:
			info = wifi_info()
			if info.get("connected") and info.get("ssid") == ssid:
				if expected_band.replace(" ", "").upper() == "MLO":
					return True
				actual_band = info.get("band", "").replace(" ", "").lower()
				return actual_band == expected_band.replace(" ", "").lower()
			time.sleep(poll_interval)
			waited += poll_interval
		return False

	if connect_and_check(timeout):
		return True

	# Wrong band (typically one SSID shared by every band). First try pinning
	# the connection to that band's radio; no Administrator rights needed.
	if expected_band.replace(" ", "").lower() in _BAND_PREF_PATTERNS:
		if connect_pinned_to_band(ssid, expected_band):
			return True
		# Fallback: ask the adapter driver to prefer the band (needs Administrator).
		previous = _set_band_preference(band=expected_band)
		if previous is None:
			print(f"  Could not reach {expected_band} on '{ssid}': {_switch_detail or 'reason unknown'}.")
			return False
		if "value" not in _original_band_pref:
			_original_band_pref["value"] = previous
			atexit.register(restore_band_preference)
		time.sleep(5)
		subprocess.run(["netsh", "wlan", "disconnect"], capture_output=True, text=True, timeout=10, check=False)
		time.sleep(2)
		return connect_and_check(40)
	return False


def active_connection_type():
	if platform.system() != "Windows":
		return "unknown"
	default_adapter = powershell(
		"$r=Get-NetRoute -DestinationPrefix '0.0.0.0/0' -AddressFamily IPv4 | "
		"Sort-Object RouteMetric,InterfaceMetric; foreach($x in $r){$n=Get-NetAdapter "
		"-InterfaceIndex $x.InterfaceIndex -ErrorAction SilentlyContinue; if($n.Status -eq 'Up'){"
		"\"$($n.Name)|$($n.InterfaceDescription)\"; break}}"
	).lower()
	if any(word in default_adapter for word in ["wi-fi", "wireless", "wlan"]):
		return "wireless"
	if default_adapter.strip():
		return "wired"
	if wifi_info()["connected"]:
		return "wireless"
	adapters = powershell("Get-NetAdapter | ? Status -eq 'Up' | select -ExpandProperty InterfaceDescription").lower()
	if adapters and not any(word in adapters for word in ["wi-fi", "wireless", "wlan"]):
		return "wired"
	return "unknown"


def client_link_info(ip):
	"""
	(adapter name, negotiated speed in Mbps) of this PC's network adapter that
	carries `ip`, or ("", None) if it can't be read. A wired link that quietly
	negotiated 100 Mbps (bad cable or port, or a forced adapter setting) caps every
	transfer near 11 MB/s no matter what the router or the drive can do.
	"""
	if platform.system() != "Windows" or not ip:
		return "", None
	script = (
		"$i=(Get-NetIPAddress -IPAddress '" + ip + "' -AddressFamily IPv4 -ErrorAction SilentlyContinue "
		"| Select-Object -First 1).InterfaceIndex; "
		"if($i){$a=Get-NetAdapter -InterfaceIndex $i; $a.Name + '|' + $a.Speed}"
	)
	line = (powershell(script, timeout=10).strip().splitlines() or [""])[-1]
	name, _, speed = line.partition("|")
	try:
		return name.strip(), float(speed) / 1_000_000
	except ValueError:
		return "", None


def dut_ip(config):
	ip = config.get("dut", {}).get("ip_address", "")
	if ip and not is_unknown(ip):
		return ip
	return share_host(config.get("usb_share_path", ""))


def ping_host(ip):
	if not ip:
		return False
	result = subprocess.run(["ping", "-n", "1", "-w", "1000", ip], capture_output=True, text=True, check=False)
	return result.returncode == 0


def validate_share_write_access(share, timeout=8):
	"""
	Check whether a share is reachable and writable, bounded by a timeout.
	os.path.isdir()/open() on a network (UNC) path are blocking OS-level
	calls with no native timeout -- on a flaky connection (common on
	wireless) they can hang indefinitely with zero output, silently
	freezing the whole run with no way to tell it's stuck. Running the
	check in a background thread with a join timeout prevents that.
	"""
	if not share:
		return False

	result = {"ok": False}

	def _check():
		try:
			if not os.path.isdir(share):
				return
			probe = os.path.join(share, f".__usb_write_probe_{os.getpid()}.tmp")
			with open(probe, "wb") as file:
				file.write(b"ok")
			os.remove(probe)
			result["ok"] = True
		except OSError:
			pass

	worker = threading.Thread(target=_check, daemon=True)
	worker.start()
	worker.join(timeout)
	if worker.is_alive():
		print(f"Warning: {share} did not respond within {timeout}s -- treating as unreachable.")
		return False
	return result["ok"]


def isdir_bounded(path, timeout=5):
	"""os.path.isdir() with a time limit. On a network path it can block for a
	long time (unreachable host, slow name lookup) with no timeout of its own."""
	result = {"ok": False}

	def _check():
		try:
			result["ok"] = os.path.isdir(path)
		except OSError:
			pass

	worker = threading.Thread(target=_check, daemon=True)
	worker.start()
	worker.join(timeout)
	return result["ok"]


def wait_for_share(share, total_timeout=45):
	"""
	Wait for a share to become reachable, printing progress. Right after a
	WiFi band switch the link is up before DHCP finishes and before Windows
	re-establishes its SMB session, so a single 8-second check often fails
	on a network that is fine a few seconds later -- which skipped whole
	bands. Retry until it responds or total_timeout passes.
	"""
	start = time.perf_counter()
	attempt = 0
	while True:
		attempt += 1
		if validate_share_write_access(share, timeout=6):
			return True
		elapsed = time.perf_counter() - start
		if elapsed >= total_timeout:
			return False
		print(f"  Waiting for {share} to become reachable ({elapsed:.0f}s of {total_timeout}s)...")
		time.sleep(2)


def get_volume_filesystem_native(share):
	"""
	Query the real filesystem type of a network share directly via the
	Windows GetVolumeInformationW API -- the same mechanism Windows itself
	uses to show "File system: NTFS" in a folder's Properties dialog.
	Unlike fsutil (unreliable on UNC paths to non-Windows SMB servers) or
	Get-SmbMapping (only works if the share is already mapped to a drive
	letter), this works directly on a UNC path like \\\\USB1\\G with no
	drive-letter mapping required. Windows-only; returns None elsewhere
	or if the call fails for any reason.
	"""
	if platform.system() != "Windows":
		return None
	try:
		path = share if share.endswith("\\") else share + "\\"
		fs_name_buffer = ctypes.create_unicode_buffer(256)
		volume_name_buffer = ctypes.create_unicode_buffer(256)
		max_component_len = ctypes.c_ulong(0)
		fs_flags = ctypes.c_ulong(0)
		serial_number = ctypes.c_ulong(0)
		success = ctypes.windll.kernel32.GetVolumeInformationW(
			ctypes.c_wchar_p(path),
			volume_name_buffer, ctypes.sizeof(volume_name_buffer),
			ctypes.byref(serial_number),
			ctypes.byref(max_component_len),
			ctypes.byref(fs_flags),
			fs_name_buffer, ctypes.sizeof(fs_name_buffer),
		)
		if success and fs_name_buffer.value:
			return fs_name_buffer.value
	except Exception:
		pass
	return None


_FORMAT_CACHE = {}


def detect_usb_format(config):
	share = config.get("usb_share_path", "")
	if not share:
		return "Unknown"
	# Measured from the drive first (see fingerprint_share_format): the name Windows
	# reports for a router share is the same for every partition and was showing NTFS
	# for an exFAT drive. Once per share -- this is called for every test round.
	if share not in _FORMAT_CACHE:
		_FORMAT_CACHE[share] = fingerprint_share_format(share) or _detect_usb_format_from_os(share)
	return _FORMAT_CACHE[share]


def _detect_usb_format_from_os(share):
	# Method 1: native Windows API -- most reliable for UNC network paths,
	# works without needing the share mapped to a drive letter first.
	native_result = get_volume_filesystem_native(share)
	if native_result:
		return native_result
	# Spawning a console app (cmd/powershell) from a process with no console
	# of its own -- which is exactly the situation here, since the GUI
	# launches usb_test.py via subprocess.Popen without one -- can hang at
	# process creation on Windows, before subprocess.run's timeout= even
	# starts counting (the hang happens before Python's tracked process
	# object exists). CREATE_NO_WINDOW avoids the console-allocation step
	# entirely instead of trying to time-bound something timeout doesn't
	# actually cover.
	no_window = subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0
	# Method 2: fsutil on UNC path (works on some Windows configurations)
	try:
		result = subprocess.run(
			["cmd", "/c", "fsutil", "fsinfo", "volumeinfo", share],
			capture_output=True, text=True, timeout=5, check=False,
			creationflags=no_window,
		)
		text = result.stdout + result.stderr
		match = re.search(r"File System Name\s*:\s*(\S+)", text, re.I)
		if match:
			return match.group(1)
	except (OSError, subprocess.TimeoutExpired):
		pass
	# Method 3: PowerShell — check if share is already mapped to a drive letter
	try:
		result = subprocess.run(
			["powershell", "-NoProfile", "-Command",
			 f"$m = Get-SmbMapping | Where-Object {{ $_.RemotePath -eq '{share}' }}; "
			 "if ($m) { (Get-Volume ($m.LocalPath.TrimEnd(':'))).FileSystem }"],
			capture_output=True, text=True, timeout=5, check=False,
			creationflags=no_window,
		)
		fs = result.stdout.strip()
		if fs and fs.lower() not in ("", "none"):
			return fs
	except (OSError, subprocess.TimeoutExpired):
		pass
	return "Unknown"


def active_usb_formats(config):
	current = config.get("current_usb_format", "").strip()
	if config.get("test_all_configured_usb_formats", False):
		return config.get("usb_formats", [])
	if current and current.lower() != "auto":
		return [current]
	if current.lower() == "auto":
		# Windows reports the same filesystem name for every partition behind a
		# router, so measure it from the drive first; fall back to what the OS says.
		detected = detect_usb_format(config)
		if is_unknown(detected):
			if config.get("prompt_for_missing_info", True):
				configured = [str(item) for item in config.get("usb_formats", [])]
				base = ["FAT32", "NTFS", "exFAT"]
				choices = base + [f for f in configured if f.upper() not in {c.upper() for c in base}]
				detected = prompt_choice("USB filesystem format", choices + ["Unknown"], "") or "Unknown"
			else:
				print(f"USB format for {config.get('usb_share_path', '')} could not be auto-detected; recording as Unknown.")
		return [detected]
	if config.get("prompt_for_missing_info", True):
		configured = [str(item) for item in config.get("usb_formats", [])]
		if configured:
			return [prompt_choice("Current USB filesystem format", configured + ["Unknown"], "") or "Unknown"]
	return ["Unknown"]


def show_progress(done, total, label, started=None):
	width = 28
	filled = int(width * done / max(total, 1))
	bar = "#" * filled + "-" * (width - filled)
	eta = ""
	if started and done:
		elapsed = time.perf_counter() - started
		remaining = elapsed / done * max(total - done, 0)
		eta = f" ETA {remaining / 60:.1f} min"
	print(f"\r[{bar}] {done}/{total} {label[:58]:58}{eta:>14}", end="", flush=True)


def show_matrix_summary(config, interfaces):
	files = config.get("test_files", [])
	# When not running the full format matrix, only one format is tested per run
	formats_count = (
		len(config.get("usb_formats", []))
		if config.get("test_all_configured_usb_formats", False)
		else 1
	)
	iterations = int(config.get("iterations", 3))
	rounds = len(interfaces) * formats_count * len(files) * 2 * iterations
	data_mb = len(interfaces) * formats_count * iterations * 2 * sum(int(item["size_mb"]) for item in files)
	print(f"Planned rounds: {rounds}")
	print(f"Approx data copied: {data_mb / 1024:.1f} GB")
	if data_mb > 20 * 1024:
		print("Warning: this is a large run. Use fewer interfaces, formats, files, or iterations for a smoke test.")


MAX_COPY_SECONDS = 3600

# The GUI's Stop button creates this file next to the program; the engine looks
# for it between rounds and during copies, then stops cleanly and saves
# whatever was measured so far.
STOP_FLAG = "stop_requested.flag"
_PROCESS_START = time.time()


class StopRequested(BaseException):
	"""Raised inside a copy when Stop was pressed. BaseException on purpose:
	run_case() catches Exception and would otherwise record it as an ERROR row."""


def stop_requested():
	try:
		return os.path.exists(resolve_path(STOP_FLAG))
	except OSError:
		return False


def _wait_until_writable(path, seconds=8.0):
	"""Wait until `path` can be created/opened for writing. Returns quietly either
	way: if it never becomes writable, the real copy then reports the real error."""
	end = time.time() + seconds
	while True:
		try:
			with open(path, "ab"):
				return
		except PermissionError:
			if time.time() >= end:
				return
			time.sleep(0.1)
		except OSError:
			return


def copy_timed(src, dst, timeout, label="", chunk_size=1 * MB):
	# timeout is an inactivity limit: it trips only when a single chunk takes
	# longer than `timeout`, so a slow-but-progressing copy is never killed.
	size = os.path.getsize(src)
	start = time.perf_counter()
	last_progress = start
	copied = 0
	last_step = 0
	chunks_done = 0
	ensure_dir(os.path.dirname(dst) or ".")
	with open(src, "rb") as source, open(dst, "wb") as target:
		while True:
			now = time.perf_counter()
			if now - last_progress > timeout:
				raise TimeoutError(f"copy stalled: no progress for {timeout} seconds at {copied * 100 / max(size, 1):.0f}%")
			if now - start > MAX_COPY_SECONDS:
				raise TimeoutError(f"copy exceeded {MAX_COPY_SECONDS} seconds total")
			chunk = source.read(chunk_size)
			if not chunk:
				break
			target.write(chunk)
			last_progress = time.perf_counter()
			copied += len(chunk)
			chunks_done += 1
			if chunks_done % 8 == 0 and stop_requested():
				raise StopRequested()
			if label:
				step = int(copied * 10 // max(size, 1))
				if step > last_step:
					last_step = step
					print(f"\r{label[:55]:55} {step * 10:6.1f}%", end="", flush=True)
		target.flush()
		os.fsync(target.fileno())
	duration = max(time.perf_counter() - start, 0.001)
	return size / MB / duration, duration


def make_test_file(path, size_mb):
	if os.path.exists(path) and os.path.getsize(path) == size_mb * MB:
		return
	ensure_dir(os.path.dirname(path) or ".")
	with open(path, "wb") as test_file:
		block = os.urandom(MB)
		for _ in range(size_mb):
			test_file.write(block)


def normalize_interfaces(items):
	normalized = []
	for item in items:
		if isinstance(item, dict):
			normalized.append(item)
		else:
			kind = "wireless" if any(x in item.lower() for x in ["ghz", "mlo", "ssid"]) else "wired"
			normalized.append({"name": item, "connection_type": kind})
	return normalized


def filter_interfaces(interfaces, active_type, wifi):
	if active_type in ["wired", "wireless"]:
		interfaces = [item for item in interfaces if item.get("connection_type") == active_type]
	if active_type == "wireless" and wifi.get("band"):
		band = wifi["band"].replace(" ", "").lower()
		interfaces = [item for item in interfaces if item.get("connection_type") != "wireless" or item.get("name", "").replace(" ", "").lower() == band]
	return interfaces


def threshold(config, direction, connection_type, size_mb=None):
	"""
	Minimum MB/s for a round to pass. A per-file-size table
	(thresholds_by_file_mbps: {direction: {connection: {"1": 21, "10": 53, ...}}})
	is used when it covers this connection -- small files are slower because
	per-file overhead dominates, so one flat number is unfair to them. A file
	between two listed sizes uses the next size down; one smaller than every
	listed size uses the smallest. Otherwise the flat thresholds_mbps value applies.
	"""
	values = config.get("thresholds_mbps", {})
	flat = values.get(direction, {}).get(connection_type, config.get("pass_threshold_mbps", 0))
	table = (config.get("thresholds_by_file_mbps") or {}).get(direction, {}).get(connection_type)
	if table and size_mb is not None:
		try:
			sizes = sorted((float(k), float(v)) for k, v in table.items())
			at_or_below = [v for k, v in sizes if k <= float(size_mb)]
			return at_or_below[-1] if at_or_below else sizes[0][1]
		except (TypeError, ValueError):
			return flat
	return flat


def threshold_lines(config):
	"""Human-readable pass-threshold description, one (label, text) per line."""
	flat = config.get("thresholds_mbps") or {}
	table = config.get("thresholds_by_file_mbps") or {}
	lines = []
	for connection in ("wired", "wireless"):
		by_size = {d: {float(k): float(v) for k, v in ((table.get(d) or {}).get(connection) or {}).items()}
		           for d in ("read", "write")}
		if any(by_size.values()):
			sizes = sorted({s for per in by_size.values() for s in per})
			head = " / ".join(f"{s:g}" for s in sizes)

			def row(direction):
				return " / ".join(f"{by_size[direction][s]:g}" if s in by_size[direction] else "-" for s in sizes)
			lines.append((f"Pass thresholds ({connection})",
			              f"By file size ({head} MB):  read ≥ {row('read')}  ·  write ≥ {row('write')} MB/s"))
		else:
			r, w = (flat.get("read") or {}).get(connection), (flat.get("write") or {}).get(connection)
			if r is not None or w is not None:
				lines.append((f"Pass thresholds ({connection})", f"Read ≥ {r:g}  ·  Write ≥ {w:g} MB/s"))
	note = config.get("thresholds_note")
	if note:
		lines.append(("Threshold basis", str(note)))
	return lines


def base_result(config, interface, usb_format, direction, test_file, iteration):
	dut = config.get("dut", {})
	client = config.get("client", {})
	# Query live, not config.get("actual_wifi", {}) -- that value is a
	# snapshot taken once at the very start of run(), before any wireless
	# band-switching happens. Using it here would silently mislabel every
	# round with whatever band the run started on, regardless of which
	# band was actually being tested at the time.
	wifi = wifi_info() if interface.get("connection_type") == "wireless" else config.get("actual_wifi", {})

	detected_band = wifi.get("band", "")
	expected_interface = interface.get("name", "unknown")

	interface_match = True

	if "ghz" in expected_interface.lower():

		interface_match = (
			expected_interface.replace(" ", "").lower()
			==
			detected_band.replace(" ", "").lower()
		)

	return Result(
		project_name=config.get("project_name", "USB Performance Test"),
		dut_model=dut.get("model", "Unknown"),
		dut_role=dut.get("role", "Unknown"),
		role_assumed=dut.get("role_assumed", False),
		dut_name=dut.get("name", ""),
		firmware_version=dut.get("firmware_version", "Unknown"),
		topology=config.get("topology", "Unknown"),
		client_type=client.get("type") or detect_client_type(),
		client_os=client.get("os") or detect_client_os(),
		client_ip=client.get("ip_address") or client_ip(),
		configured_interface=expected_interface,
		connection_type=interface.get("connection_type", "unknown"),
		actual_ssid=wifi.get("ssid", ""),
		actual_channel=wifi.get("channel", ""),
		actual_band=detected_band,
		detected_gateway_ip=config.get("detected_gateway_ip", ""),
		detected_hardware_version=dut.get("hardware_version", ""),
		bssid=wifi.get("bssid", ""),
		wifi_signal_percent=wifi.get("signal", ""),
		usb_share_path=config.get("usb_share_path", ""),
		actual_usb_format=detect_usb_format(config),
		expected_interface=expected_interface,
		interface_match=interface_match,
		dut_reachable=config.get("dut_reachable", False),
		share_write_access=config.get("share_write_access", False),
		usb_format=usb_format,
		direction=direction,
		file_set_version=config.get("test_file_set_version", "unversioned"),
		file_name=test_file["name"],
		iteration=iteration,
		throughput_mbps=0.0,
		duration_seconds=0.0,
		timestamp=datetime.now().isoformat(timespec="seconds"),
		file_verification=test_file.get("verification", ""),
	)


def skipped_result(config, interface_name, reason):
	"""A row for something that could not be tested, so the report shows it."""
	dut = config.get("dut", {})
	client = config.get("client", {})
	kind = "wireless" if any(x in interface_name.lower() for x in ["ghz", "mlo", "ssid"]) else "wired"
	return Result(
		project_name=config.get("project_name", "USB Performance Test"),
		dut_model=dut.get("model", "Unknown"),
		dut_role=dut.get("role", "Unknown"),
		firmware_version=dut.get("firmware_version", "Unknown"),
		topology=config.get("topology", "Unknown"),
		client_type=client.get("type") or detect_client_type(),
		client_os=client.get("os") or detect_client_os(),
		client_ip=client.get("ip_address") or client_ip(),
		configured_interface=interface_name,
		connection_type=kind,
		actual_ssid="",
		actual_channel="",
		actual_band="",
		detected_gateway_ip=config.get("detected_gateway_ip", ""),
		dut_name=dut.get("name", ""),
		detected_hardware_version=dut.get("hardware_version", ""),
		usb_share_path=config.get("usb_share_path", ""),
		expected_interface=interface_name,
		interface_match=True,
		role_assumed=dut.get("role_assumed", False),
		dut_reachable=config.get("dut_reachable", False),
		share_write_access=config.get("share_write_access", False),
		file_set_version=config.get("test_file_set_version", "unversioned"),
		status="SKIPPED",
		failure_reason=reason,
		timestamp=datetime.now().isoformat(timespec="seconds"),
	)


def run_case(config, interface, usb_format, direction, test_file, iteration, label=""):
	result = base_result(config, interface, usb_format, direction, test_file, iteration)
	local_dir = resolve_path(config.get("local_test_file_dir", "test_files"))
	output_dir = resolve_path(config.get("output_path", "test_results"))
	share = config.get("usb_share_path", "")
	timeout = config.get("test_timeout_seconds", 180)
	local_src = os.path.join(local_dir, test_file["name"])
	usb_file = os.path.join(share, test_file["name"])
	local_read = os.path.join(output_dir, "readback", f"{iteration}_{test_file['name']}")

	try:
		if not share or not os.path.isdir(share):
			raise FileNotFoundError(f"USB share is not mounted or reachable: {share}")
		if test_file.get("generate", True):
			make_test_file(local_src, int(test_file["size_mb"]))
		if direction == "write":
			free_bytes = shutil.disk_usage(share).free
			file_bytes = os.path.getsize(local_src) if os.path.exists(local_src) else int(test_file.get("size_mb", 0)) * MB
			if free_bytes < file_bytes:
				raise OSError(
					f"Insufficient USB storage: {free_bytes // MB} MB free, need {file_bytes // MB} MB"
				)
			# The previous round deleted this same file name a moment ago; over SMB it can
			# still be "delete pending", and creating it again fails with Permission denied.
			# Wait here, before the clock starts, so the measurement is unaffected.
			_wait_until_writable(usb_file)
			speed, duration = copy_timed(local_src, usb_file, timeout, label)
		else:
			# Fresh remote file per iteration so the SMB client can't serve a
			# repeat read from its local cache (see KNOWN_LIMITATIONS 17).
			# This staging copy is untimed (not part of the measured result) but
			# still needs the same stall protection as every other share write --
			# a bare shutil.copyfile() here has no timeout and can hang forever
			# if the share stops responding mid-write.
			read_src = os.path.join(share, f"_read_{iteration}_{test_file['name']}")
			try:
				# The staging copy sits inside the cleanup too: when it is interrupted (Stop)
				# or fails part-way, its half-written file must not be left on the drive.
				copy_timed(local_src, read_src, timeout, label="")
				speed, duration = copy_timed(read_src, local_read, timeout, label)
			finally:
				try:
					os.remove(read_src)
				except OSError:
					pass
		result.throughput_mbps = round(speed, 2)
		result.duration_seconds = round(duration, 3)
		# Nothing here can move data faster than a 10 Gbps link (~1250 MB/s);
		# a result above that is a local/cached copy, not a real measurement.
		ceiling = float(config.get("max_plausible_mbps", 1250))
		if speed <= 0 or speed > ceiling:
			raise ValueError(
				f"Abnormal throughput result ({speed:.2f} MB/s; expected between 0 and {ceiling:.0f}) -- measurement discarded"
			)
		limit = threshold(config, direction, result.connection_type, test_file.get("size_mb"))
		result.passed = speed >= limit
		result.status = "PASS" if result.passed else "PERF_FAIL"
		if not result.passed:
			result.failure_reason = f"throughput below configured threshold ({speed:.2f} < {limit:.2f} MB/s)"
	except Exception as error:
		result.status = "ERROR"
		result.failure_reason = type(error).__name__ + ": " + str(error)
	finally:
		if direction == "read" and os.path.exists(local_read):
			try:
				os.remove(local_read)
			except OSError:
				pass
		# The test file the write round put on the USB drive is the tool's own
		# copy; don't leave ~660 MB of test files on the customer's drive.
		# Untimed: the measurement above already ended (after fsync).
		if direction == "write":
			try:
				os.remove(usb_file)
			except OSError:
				pass
	return result


def add_stats(results):
	groups = {}
	for result in results:
		key = (result.dut_role, result.configured_interface, result.usb_format, result.direction, result.file_name)
		groups.setdefault(key, []).append(result)
	for group in groups.values():
		speeds = [item.throughput_mbps for item in group if item.duration_seconds > 0]
		if not speeds:
			continue
		avg = statistics.mean(speeds)
		stddev = statistics.stdev(speeds) if len(speeds) > 1 else 0.0
		for item in group:
			item.average_mbps = round(avg, 2)
			item.min_mbps = round(min(speeds), 2)
			item.max_mbps = round(max(speeds), 2)
			item.stddev_mbps = round(stddev, 2)


def summary_rows(results, *fields):
	groups = {}
	for result in results:
		key = tuple(getattr(result, field) for field in fields)
		groups.setdefault(key, []).append(result)
	rows = ""
	for key, items in sorted(groups.items()):
		speeds = [item.throughput_mbps for item in items if item.duration_seconds > 0]
		avg = statistics.mean(speeds) if speeds else 0
		passed = sum(1 for item in items if item.passed)
		label = " / ".join(str(part) for part in key)
		rows += f"<tr><td>{label}</td><td>{len(items)}</td><td>{passed}</td><td>{avg:.2f}</td></tr>"
	return rows


def results_from_dicts(rows):
	"""Reconstruct Result dataclass objects from serialized dict rows."""
	fields = set(Result.__dataclass_fields__)
	out = []
	for row in rows:
		try:
			out.append(Result(**{k: v for k, v in row.items() if k in fields}))
		except TypeError:
			pass
	return out


def save_outputs(config, results):
	out_dir = resolve_path(config.get("output_path", "test_results"))
	ensure_dir(out_dir)

	if config.get("accumulate_results", False):
		# Archive this run as an individual timestamped file
		runs_dir = os.path.join(out_dir, "runs")
		ensure_dir(runs_dir)
		# Seconds included, and never an existing name: with minute precision two runs
		# started within the same minute silently overwrote each other, losing one.
		ts = (
			results[0].timestamp.replace(":", "").replace("T", "_")[:17]
			if results else datetime.now().strftime("%Y%m%d_%H%M%S")
		)
		run_file, suffix = os.path.join(runs_dir, f"run_{ts}.json"), 1
		while os.path.exists(run_file):
			suffix += 1
			run_file = os.path.join(runs_dir, f"run_{ts}_{suffix}.json")
		with open(run_file, "w", encoding="utf-8") as f:
			json.dump([asdict(r) for r in results], f, indent=2)
		# Merge every archived run into one combined result set
		all_rows = []
		for fname in sorted(os.listdir(runs_dir)):
			if fname.endswith(".json"):
				try:
					with open(os.path.join(runs_dir, fname), encoding="utf-8") as f:
						all_rows.extend(json.load(f))
				except Exception:
					pass
		combined = results_from_dicts(all_rows)
		if len(combined) > len(results):
			formats = sorted({r.usb_format for r in combined})
			print(f"Accumulated results across {len(formats)} format(s): {', '.join(formats)}")
	else:
		combined = list(results)

	add_stats(combined)
	rows = [asdict(r) for r in combined]
	with open(os.path.join(out_dir, "test_results.json"), "w", encoding="utf-8") as json_file:
		json.dump(rows, json_file, indent=2)
	with open(os.path.join(out_dir, "test_results.csv"), "w", newline="", encoding="utf-8") as csv_file:
		fieldnames = rows[0].keys() if rows else Result.__dataclass_fields__.keys()
		writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
		writer.writeheader()
		writer.writerows(rows)
	save_html(config, combined, os.path.join(out_dir, "report.html"))
	save_excel(config, combined, out_dir)


def save_excel(config, results, out_dir):
	"""Formatted Excel workbook plus the blank template. Never allowed to fail the run: the CSV, JSON and HTML are already written."""
	if not results:
		return
	try:
		import excel_report
		path = excel_report.write_workbook(out_dir, results, dict(config, log_file=os.path.basename(LOG_PATH or "")))
		excel_report.write_template(out_dir)
		print(f"Excel workbook: {path}")
		logging.info(f"Excel workbook written: {path}")
	except ImportError:
		print("Excel workbook skipped: the 'xlsxwriter' package is not installed (pip install xlsxwriter).")
	except Exception as error:
		print(f"Excel workbook could not be written ({type(error).__name__}: {error}). The CSV, JSON and HTML reports are unaffected.")
		logging.warning(f"Excel workbook failed: {type(error).__name__}: {error}")


def build_chart_data(results):
	"""Aggregate results into chart-ready data structures."""
	def avg(subset):
		speeds = [r.throughput_mbps for r in subset if r.duration_seconds > 0]
		return round(statistics.mean(speeds), 2) if speeds else 0
	interfaces = sorted({r.configured_interface for r in results})
	formats = sorted({r.usb_format for r in results})
	roles = sorted({r.dut_role for r in results if r.dut_role not in ("Unknown", "")})
	conn_types = ["wired", "wireless"]
	conn_read = [
			avg([ r for r in results
			if r.connection_type == ct and r.direction == "read"])
			for ct in conn_types
	]
	conn_write = [
		avg([ r for r in results
	   		if r.connection_type == ct and r.direction == "write"
		])
		for ct in conn_types
	]
	per_round = []
	for iface in interfaces:
		for direction in ["read", "write"]:
			data = [r.throughput_mbps for r in results
			        if r.configured_interface == iface and r.direction == direction and r.duration_seconds > 0]
			if data:
				per_round.append({"label": f"{iface} {direction}", "data": data})
	max_rounds = max((len(d["data"]) for d in per_round), default=0)
	return {
		"interfaces": interfaces,
		"iface_read": [avg([r for r in results if r.configured_interface == i and r.direction == "read"]) for i in interfaces],
		"iface_write": [avg([r for r in results if r.configured_interface == i and r.direction == "write"]) for i in interfaces],
		"formats": formats,
		"fmt_read": [avg([r for r in results if r.usb_format == f and r.direction == "read"]) for f in formats],
		"fmt_write": [avg([r for r in results if r.usb_format == f and r.direction == "write"]) for f in formats],
		"roles": roles,
		"role_read": [avg([r for r in results if r.dut_role == role and r.direction == "read"]) for role in roles],
		"role_write": [avg([r for r in results if r.dut_role == role and r.direction == "write"]) for role in roles],
		"show_role_chart": len(roles) > 1,
		"per_round": per_round,
		"round_labels": [f"Round {i + 1}" for i in range(max_rounds)],
		"connection_types": conn_types,
		"conn_read": conn_read,
		"conn_write": conn_write
	}


def _chart_js_tag():
	local = os.path.join(app_dir(), "chart.umd.min.js")
	try:
		with open(local, "r", encoding="utf-8") as f:
			return "<script>" + f.read().replace("</script", "<\\/script") + "</script>"
	except OSError:
		return '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>'


def _attention_reason(r):
	if r.failure_reason:
		return r.failure_reason
	return "Connected band " + (r.actual_band or "?") + " does not match configured interface " + r.configured_interface


def _environment_html(config, dut, first, results=()):
	"""
	The report's Test Environment section as three small cards (device under
	test, test setup, client PC) instead of one long table. Every value is read
	from the run itself; anything that could not be read shows as a muted
	"Unknown" rather than a guess.
	"""
	def text(value):
		value = "" if value is None else str(value).strip()
		return esc(value) if value and value.lower() != "unknown" else '<span class="na">Unknown</span>'

	def yes_no(value):
		if value is True:
			return '<span class="yes">&#10004; Yes</span>'
		if value is False:
			return '<span class="no">&#10008; No</span>'
		return '<span class="na">Not checked</span>'

	def card(title, rows):
		body = "".join(f"<dt>{esc(label)}</dt><dd>{value}</dd>" for label, value in rows)
		return f'<div class="ecard"><h3>{esc(title)}</h3><dl>{body}</dl></div>'

	name_html = text(dut.get("name"))

	try:
		when = datetime.fromisoformat(first.timestamp).strftime("%b %d, %Y  %I:%M %p")
	except (ValueError, TypeError):
		when = first.timestamp
	connection = str(config.get("actual_connection_type", "")).strip().capitalize()

	os_html = text(first.client_os)
	if platform.system() == "Windows":
		try:
			os_html += f"<small>Windows build {sys.getwindowsversion().build}, read from this PC</small>"
		except (AttributeError, OSError):
			pass

	role = str(dut.get("role", "")).strip()
	role_html = f'<span class="role">{esc(role)}</span>' if role and role.lower() != "unknown" else text(role)
	if dut.get("role_assumed"):
		role_html += "<small>assumed &mdash; could not be confirmed</small>"

	# A report merged from several runs (Router then Extender) covers more than one device:
	# list each one from its own results instead of describing only the last run.
	per_role = {}
	for item in results:
		if item.status != "SKIPPED" and item.dut_role and item.dut_role not in per_role:
			per_role[item.dut_role] = item
	if len(per_role) > 1:
		def one_line(item):
			parts = [item.dut_model, f"HW {item.detected_hardware_version}" if item.detected_hardware_version else "",
			         item.firmware_version, item.usb_share_path]
			return f'<span class="long">{esc(item.dut_name or "Unknown")}<small>' + esc("  ·  ".join(p for p in parts if p)) + "</small></span>"
		device = card("Devices under test", [
			(role, one_line(item)) for role, item in sorted(per_role.items(), key=lambda kv: (kv[0] != "Router", kv[0]))
		] + [("Gateway IP", text(config.get("detected_gateway_ip")))])
	else:
		device = card("Device under test", [
			("Name", name_html),
			("Role", role_html),
			("Model", text(dut.get("model"))),
			("Hardware version", text(dut.get("hardware_version"))),
			("Firmware", text(dut.get("firmware_version"))),
			("Gateway IP", text(config.get("detected_gateway_ip"))),
			("Reachable", yes_no(config.get("dut_reachable"))),
			("USB write access", yes_no(config.get("share_write_access"))),
		])
	topologies = list(dict.fromkeys(i.topology for i in results if i.topology and i.status != "SKIPPED")) or [config.get("topology")]
	topology_html = "<br>".join(text(t) for t in topologies)
	setup = card("Test setup", [
		("Project", text(config.get("project_name", "USB Performance Test"))),
		("Date", text(when)),
		("Topology", topology_html),
		("Connection", text(connection)),
		("Test files", '<span class="long">' + text(config.get("file_set_summary", "not verified")) + "</span>"),
		("Pass thresholds", '<span class="long">' + "<br>".join(
			f"{esc(label.replace('Pass thresholds ', '').strip('()').capitalize() or 'All')}: {esc(value)}"
			for label, value in threshold_lines(config) if not label.startswith("Threshold basis")) + "</span>"),
	])
	link = config.get("client_link") or {}
	link_html = (f"{esc(link.get('adapter', ''))}<small>negotiated {link['mbps']:g} Mbps</small>"
	             if link.get("mbps") else '<span class="na">Unknown</span>')
	client = card("Client PC", [
		("Operating system", os_html),
		("Device type", text(first.client_type)),
		("IP address", text(first.client_ip)),
		("Network link", link_html),
	])
	warning = config.get("client_link_warning")
	callout = f'<div class="callout"><b>&#9888; Read these results with care.</b> {esc(warning)}</div>\n' if warning else ""
	note = config.get("client_link_note")
	if note:
		callout += f'<div class="callout info"><b>&#8505; Good to know.</b> {esc(note)}</div>\n'
	source = {"upnp": "were read from the device itself over UPnP",
	          "config": "come from the details entered in the configuration"}.get(dut.get("info_source"), "come from the configuration")
	return (f'{callout}<div class="env">{device}{setup}{client}</div>\n'
	        f'<p class="envnote">Device details {source}. The client PC details are read from this computer. '
	        f'Anything that could not be read is shown as Unknown; nothing is assumed.</p>')


def save_html(config, results, path):
	if not results:
		return
	every_row = results
	skipped = [r for r in every_row if r.status == "SKIPPED"]
	results = [r for r in every_row if r.status != "SKIPPED"]
	# Environment fields (date, client IP, OS) come from a row that actually
	# ran: a skipped band's row is recorded mid-switch, when the client may be
	# briefly on another network and report the wrong IP.
	first = results[0] if results else every_row[0]
	cd = build_chart_data(results)
	passed = sum(1 for r in results if r.passed)
	failed = len(results) - passed
	if failed:
		verdict, verdict_class = "FAIL", "fail-b"
	elif skipped or not results:
		verdict, verdict_class = "INCOMPLETE", "warn"
	else:
		verdict, verdict_class = "PASS", "pass"
	attention = [r for r in every_row if r.status in ("PERF_FAIL", "ERROR", "SKIPPED") or not r.interface_match]
	attention_rows = "".join(
		f"<tr class=\"{'skip' if r.status == 'SKIPPED' else 'fail'}\">"
		f"<td>{esc(r.status if r.status != 'PASS' else 'ABNORMAL')}</td><td>{esc(r.dut_role)}</td>"
		f"<td>{esc(r.configured_interface)}</td><td>{esc(r.usb_format)}</td><td>{esc(r.direction)}</td>"
		f"<td>{esc(r.file_name)}</td><td>{r.iteration or ''}</td>"
		f"<td>{esc(_attention_reason(r))}</td></tr>"
		for r in attention
	)
	attention_section = (
		"<h2>Failed, Skipped and Abnormal Cases</h2>"
		"<table><tr><th>Status</th><th>Role</th><th>Interface</th><th>Format</th><th>Direction</th>"
		"<th>File</th><th>Round</th><th>Reason</th></tr>" + attention_rows + "</table>\n"
		if attention else
		"<h2>Failed, Skipped and Abnormal Cases</h2><p>None. Every configured case ran and met its threshold.</p>\n"
	)
	dut = config.get("dut", {})
	client = config.get("client", {})
	J = json.dumps
	PALETTE = ["#1f3f73","#c9a227","#5b7db1","#8a6d3b","#2f6f5e","#b5523b","#7a869a","#d8bf6b","#4b3b2a","#9aa5b8"]
	pr_datasets = J([
		{"label": d["label"], "data": d["data"],
		 "backgroundColor": PALETTE[i % len(PALETTE)],
		 "borderColor": PALETTE[i % len(PALETTE)]}
		for i, d in enumerate(cd["per_round"])
	])
	raw_rows = "".join(
		f"<tr class=\"{'skip' if r.status == 'SKIPPED' else ('' if r.passed else 'fail')}\">"
		f"<td>{esc(r.dut_role)}</td><td>{esc(r.dut_name)}</td><td>{esc(r.usb_share_path)}</td>"
		f"<td>{esc(r.configured_interface)}</td><td>{esc(r.connection_type)}</td>"
		f"<td>{esc(r.actual_ssid)}</td><td>{esc(r.actual_band)}</td>"
		f"<td>{esc(r.usb_format)}</td><td>{esc(r.direction)}</td>"
		f"<td>{esc(r.file_name)}</td><td>{esc(r.file_verification)}</td><td>{r.iteration or ''}</td>"
		f"<td>{r.throughput_mbps:.2f}</td><td>{r.average_mbps:.2f}</td>"
		f"<td>{esc(r.status)}</td><td>{esc(r.failure_reason)}</td></tr>"
		for r in every_row
	)
	role_section = (
		"<h2>Summary by DUT Role</h2>"
		"<table><tr><th>Role</th><th>Rounds</th><th>Passed</th><th>Avg MB/s</th></tr>"
		+ summary_rows(results, "dut_role") + "</table>"
	) if cd["show_role_chart"] else ""
	role_canvas = (
		'<div class="chart-box"><h3>Router vs Extender</h3><div class="cv"><canvas id="cRole"></canvas></div></div>'
	) if cd["show_role_chart"] else ""
	role_js = (
		"mkBar('cRole'," + J(cd["roles"]) + "," + J(cd["role_read"]) + "," + J(cd["role_write"]) + ");"
	) if cd["show_role_chart"] else ""
	chart_js = "\n".join([
		"var __charts=[];",
		"function mkBar(id,labels,reads,writes){",
		"  __charts.push(new Chart(document.getElementById(id),{type:'bar',",
		"    data:{labels:labels,datasets:[",
		"      {label:'Read (MB/s)',data:reads,backgroundColor:'#1f3f73'},",
		"      {label:'Write (MB/s)',data:writes,backgroundColor:'#c9a227'}]},",
		"    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'top'}},",
		"             scales:{y:{beginAtZero:true,title:{display:true,text:'MB/s'}}}}}));",
		"}",
		"function __renderCharts(){",
		"mkBar('cIface'," + J(cd["interfaces"]) + "," + J(cd["iface_read"]) + "," + J(cd["iface_write"]) + ");",
		"mkBar('cFmt'," + J(cd["formats"]) + "," + J(cd["fmt_read"]) + "," + J(cd["fmt_write"]) + ");",
		"mkBar('cConn'," + J(cd["connection_types"]) + "," + J(cd["conn_read"]) + "," + J(cd["conn_write"]) + ");",
		role_js,
		"__charts.push(new Chart(document.getElementById('cRounds'),{type:'bar',",
		"  data:{labels:" + J(cd["round_labels"]) + ",datasets:" + pr_datasets + "},",
		"  options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'right'}},",
		"           scales:{y:{beginAtZero:true,title:{display:true,text:'MB/s'}}}}}));",
		"}",
		"__renderCharts();",
		# Browser zoom changes devicePixelRatio, but Chart.js reads DPR only at
		# chart construction -- resize() alone does not pick up a DPR-only
		# change, which is why zooming left the canvas rendered at the old
		# resolution (blurry, gridlines/bars not lined up with the axes) until
		# something else forced a full rebuild. There's no native DPR-change
		# event; the commonly used substitute (re-subscribing a matchMedia
		# resolution query) did not fire reliably in testing, so this polls
		# devicePixelRatio instead -- a plain comparison a couple times a
		# second, negligible cost, and it can't miss a change regardless of
		# what does or doesn't dispatch an event for it.
		"var __lastDPR=window.devicePixelRatio;",
		"setInterval(function(){",
		"  if(window.devicePixelRatio!==__lastDPR){",
		"    __lastDPR=window.devicePixelRatio;",
		"    __charts.forEach(function(c){ c.destroy(); });",
		"    __charts=[];",
		"    __renderCharts();",
		"  }",
		"},400);",
	])
	css = (
		"body{font-family:Segoe UI,Arial,sans-serif;font-size:14px;margin:18px;color:#222}"
		".cv{position:relative;height:280px}.tw{overflow-x:auto}"
		"h1{color:#0b1f3a;margin:0 0 14px;padding-bottom:8px;border-bottom:3px solid #c9a227}"
		"h2{color:#0b1f3a;border-bottom:2px solid #c9a227;padding-bottom:3px;margin:16px 0 8px;font-size:18px}"
		"table{border-collapse:collapse;width:100%;margin-bottom:14px;font-size:13px}"
		"td,th{border:1px solid #ccc;padding:4px 7px}th{background:#0b1f3a;color:#fff}"
		"tr:nth-child(even){background:#f7f4ea}.fail td{color:#b00020;font-weight:bold}"
		".env{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:14px;margin-bottom:14px}"
		".ecard{border:1px solid #d9d1bd;border-radius:8px;background:#fff;padding:12px 16px;"
		"box-shadow:0 1px 2px rgba(0,0,0,.05)}"
		".ecard h3{margin:0 0 8px;font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:#8a6d1b}"
		".ecard dl{display:grid;grid-template-columns:max-content 1fr;gap:6px 18px;margin:0}"
		".ecard dt{color:#667788;font-size:12px;padding-top:1px}"
		".ecard dd{margin:0;font-weight:600;overflow-wrap:anywhere}"
		".ecard small{display:block;font-weight:400;color:#8896a4;font-size:11px}"
		".role{display:inline-block;background:#0b1f3a;color:#fff;border-radius:10px;padding:1px 10px;font-size:12px}"
		".yes{color:#2b7a3b}.no{color:#b00020}.na{color:#8896a4;font-weight:400}"
		".long{font-weight:400;font-size:12.5px;line-height:1.4}"
		".callout{background:#fff4e5;border:1px solid #f0b35c;border-left:5px solid #e08a00;"
		"border-radius:6px;padding:10px 14px;margin:0 0 14px;color:#6b4200}"
		".callout.info{background:#eef1f8;border-color:#a9b8d6;border-left-color:#1f3f73;color:#0b1f3a}"
		".envnote{color:#8896a4;font-size:12px;margin:-6px 0 12px}"
		".cg{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:10px}"
		".chart-box{border:1px solid #ddd;padding:8px;border-radius:4px;background:#fafafa}"
		".chart-wide{border:1px solid #ddd;padding:8px;border-radius:4px;background:#fafafa;margin-bottom:10px}"
		".badge{display:inline-block;padding:3px 12px;border-radius:12px;font-weight:bold;margin:0 3px}"
		".pass{background:#e2efda;color:#375623}.fail-b{background:#fce4d6;color:#843c0c}"
		".warn{background:#fff2cc;color:#7f6000}.skip td{color:#7f6000;font-weight:bold}"
		"@media print{"
		"  body{margin:10px;font-size:11px}"
		"  h2{font-size:15px;margin:10px 0 6px}"
		"  .cg{display:block}"
		"  .chart-box,.chart-wide{page-break-inside:avoid;break-inside:avoid;"
		"    margin-bottom:14px;max-height:260px}"
		"  .cv{height:220px}"
		"  canvas{max-height:220px!important;max-width:100%!important}"
		"  table{page-break-inside:auto;font-size:9px}"
		"  td,th{padding:2px 4px}"
		"  tr{page-break-inside:avoid;break-inside:avoid}"
		"  h2{page-break-after:avoid;break-after:avoid}"
		# The raw-results table is wide enough to need horizontal scrolling on
		# screen (.tw{overflow-x:auto}), but a printed/PDF page can't scroll --
		# without this override, every column past the page edge was silently
		# clipped rather than shown. Shrinking its font further and letting it
		# wrap onto a landscape page (see save_report_as_pdf in usb_gui.py)
		# keeps every column visible instead.
		"  #rawWrap{overflow:visible!important}"
		"  #rawWrap table{font-size:7.5px;table-layout:auto}"
		"  #rawWrap td,#rawWrap th{padding:1.5px 3px;white-space:normal;word-break:break-word}"
		"}"
	)
	ts = first.timestamp
	client_ip_val = first.client_ip
	env_html = _environment_html(config, dut, first, results)
	html = (
		'<!doctype html><html><head><meta charset="utf-8"><title>USB Performance Report</title>\n'
		+ _chart_js_tag() + "\n"
		f"<style>{css}</style></head>\n"
		f"<body><h1>USB Performance Report</h1>\n"
		f"<h2>Test Environment</h2>\n{env_html}\n"
		f"<h2>Overall Result</h2>\n"
		f"<p><span class=\"badge {verdict_class}\">Overall: {verdict}</span>"
		f"<span class=\"badge pass\">&#10004; Passed: {passed}</span>"
		f"<span class=\"badge fail-b\">&#10008; Failed: {failed}</span>"
		f"<span class=\"badge warn\">Skipped: {len(skipped)}</span>"
		f"&nbsp; Total: {len(results)} rounds run</p>\n"
		f"{attention_section}"
		f"<h2>Performance Charts</h2>\n"
		f'<div class="cg">\n'
		f'<div class="chart-box"><h3>Avg Throughput by Interface</h3><div class="cv"><canvas id="cIface"></canvas></div></div>\n'
		f'<div class="chart-box"><h3>Avg Throughput by USB Format</h3><div class="cv"><canvas id="cFmt"></canvas></div></div>\n'
		f'<div class="chart-box"><h3>Wired vs Wireless</h3><div class="cv"><canvas id="cConn"></canvas></div></div>\n'
		f"{role_canvas}\n"
		f"</div>\n"
		f'<div class="chart-wide"><h3>Per-Round Throughput Variation</h3>'
		f'<div class="cv"><canvas id="cRounds"></canvas></div></div>\n'
		f"<h2>Summary by Direction</h2>"
		f"<table><tr><th>Direction</th><th>Rounds</th><th>Passed</th><th>Avg MB/s</th></tr>"
		f"{summary_rows(results, 'direction')}</table>\n"
		f"<h2>Summary by Interface</h2>"
		f"<table><tr><th>Interface</th><th>Rounds</th><th>Passed</th><th>Avg MB/s</th></tr>"
		f"{summary_rows(results, 'configured_interface')}</table>\n"
		f"<h2>Summary by USB Format</h2>"
		f"<table><tr><th>Format</th><th>Rounds</th><th>Passed</th><th>Avg MB/s</th></tr>"
		f"{summary_rows(results, 'usb_format')}</table>\n"
		f"<h2>Summary by Format and Direction</h2>"
		f"<table><tr><th>Format / Direction</th><th>Rounds</th><th>Passed</th><th>Avg MB/s</th></tr>"
		f"{summary_rows(results, 'usb_format', 'direction')}</table>\n"
		f"{role_section}\n"
		f"<h2>Raw Results</h2>"
		f"<div class=\"tw\" id=\"rawWrap\"><table><tr><th>Role</th><th>DUT Name</th><th>Share Path</th><th>Interface</th><th>Type</th>"
		f"<th>SSID</th><th>Band</th><th>Format</th><th>Direction</th>"
		f"<th>File</th><th>File Check</th><th>Round</th><th>MB/s</th><th>Avg MB/s</th>"
		f"<th>Result</th><th>Failure</th></tr>\n"
		f"{raw_rows}</table></div>\n"
		f"<p style=\"color:#888;font-size:12px\">Log file: {esc(LOG_PATH or os.path.abspath('logs'))}</p>\n"
		f"<script>\n{chart_js}\n</script>\n"
		"</body></html>"
	)
	with open(path, "w", encoding="utf-8") as html_file:
		html_file.write(html)



def _build_target_config(base_config, target):
	"""Build a per-target config dict for use in multi-target test runs."""
	tc = dict(base_config)
	share = target.get("usb_share_path", "") or target_share_path(base_config, target)
	if share:
		tc["usb_share_path"] = share
	# Per-target format override: some routers report the same filesystem name for
	# every share on one physical drive regardless of each partition's real format,
	# so a known (or measured) format can be set per target.
	if target.get("current_usb_format") and not is_unknown(target["current_usb_format"]):
		tc["current_usb_format"] = target["current_usb_format"]
	role = target.get("role", "auto")
	if is_unknown(role):
		role = infer_dut_role(resolve_host_ip(share_host(share)), base_config.get("detected_gateway_ip", ""))
	dut = dict(base_config.get("dut", {}))
	dut["role"] = role
	for k in ("name", "model", "firmware_version"):
		if target.get(k) and not is_unknown(target[k]):
			dut[k] = target[k]
	tc["dut"] = dut
	tc["share_write_access"] = validate_share_write_access(share)
	tc["dut_reachable"] = ping_host(resolve_host_ip(share_host(share))) if share else False
	return tc


def save_checkpoint(done):
	with open(resolve_path("checkpoint.json"), "w", encoding="utf-8") as f:
		json.dump({"completed": done}, f, indent=2)

def _list_disk_shares(host, timeout=5):
	"""Return names of non-hidden disk shares advertised by a host via net view."""
	try:
		result = subprocess.run(
			["net", "view", f"\\\\{host}", "/all"],
			capture_output=True, text=True, timeout=timeout, check=False,
		)
		names = []
		for line in result.stdout.splitlines():
			parts = line.split()
			# net view output: ShareName  Type  UsedAs  Comment
			# Filter to Disk shares; exclude hidden admin shares (IPC$, ADMIN$, C$…)
			if len(parts) >= 2 and parts[1].lower() == "disk" and not parts[0].endswith("$"):
				names.append(parts[0])
		return names
	except Exception:
		return []


def discover_partition_shares(host):
	"""
	List every writable disk share on a given host as full UNC paths.
	Used to find multiple partitions of one physical USB drive that the
	router exposes as separate shares (e.g. \\\\USB1\\g, \\\\USB1\\h,
	\\\\USB1\\i for three differently-formatted partitions on one drive).
	Unlike resolve_usb_share(), this does not scan the whole network —
	it only lists shares on the specific host given.
	"""
	shares = []
	for name in _list_disk_shares(host):
		path = f"\\\\{host}\\{name}"
		if validate_share_write_access(path):
			shares.append(path)
	return shares



def candidate_share_hosts(preferred=()):
	"""
	Addresses worth checking for USB storage, most likely first: hosts that worked
	before, then the default gateway(s) -- where a router's own USB port is shared.
	Private addresses only. Returns (hosts, trusted); every candidate is trusted,
	because nothing here is guessed from the wider network.
	"""
	import ipaddress
	hosts = []

	def add(value):
		try:
			addr = ipaddress.ip_address(str(value).strip())
		except ValueError:
			return
		if addr.version != 4 or not addr.is_private or addr.is_loopback or addr.is_link_local or addr.is_multicast:
			return
		if str(addr) not in hosts:
			hosts.append(str(addr))

	for host in preferred:
		add(host)
	script = "Get-NetRoute -DestinationPrefix '0.0.0.0/0' -AddressFamily IPv4 | ForEach-Object { $_.NextHop }"
	for line in powershell(script, timeout=10).splitlines():
		add(line)
	return hosts[:10], set(hosts)


def find_usb_shares(preferred_hosts=()):
	"""
	Find USB storage on the network with no address typed in: checks the hosts that
	worked before and the default gateway. Returns (found, needs_login): found maps
	host -> writable share paths; needs_login lists hosts that answer file sharing
	but list no shares (they want a login first).
	"""
	hosts, _ = candidate_share_hosts(preferred_hosts)
	found, needs_login = {}, []
	with ThreadPoolExecutor(max_workers=16) as pool:
		alive = [h for h, up in zip(hosts, pool.map(lambda x: host_serves_smb(x, 1.0), hosts)) if up]
		for host, names in zip(alive, pool.map(_list_disk_shares, alive)):
			if not names:
				needs_login.append(host)
				continue
			paths = [f"\\\\{host}\\{name}" for name in names]
			good = [p for p, ok in zip(paths, pool.map(validate_share_write_access, paths)) if ok]
			if good:
				found[host] = good
	return found, needs_login


def host_serves_smb(host, timeout=2):
	"""True if something answers on the Windows file-sharing port."""
	try:
		with socket.create_connection((host, 445), timeout=timeout):
			return True
	except OSError:
		return False


def sign_in_to_share_host(host, username, password):
	"""
	Log this PC in to a router's USB sharing service, the same as typing the
	credentials into File Explorer. Many routers hide their USB shares from
	a PC that hasn't signed in, so a scan finds nothing until this is done.
	The login lives in this Windows session only (the password is never
	saved by this tool). Returns (ok, message).
	"""
	if platform.system() != "Windows":
		return False, "Signing in is only supported on Windows."

	class NETRESOURCE(ctypes.Structure):
		_fields_ = [
			("dwScope", ctypes.c_ulong), ("dwType", ctypes.c_ulong),
			("dwDisplayType", ctypes.c_ulong), ("dwUsage", ctypes.c_ulong),
			("lpLocalName", ctypes.c_wchar_p), ("lpRemoteName", ctypes.c_wchar_p),
			("lpComment", ctypes.c_wchar_p), ("lpProvider", ctypes.c_wchar_p),
		]

	remote = f"\\\\{host}\\IPC$"
	resource = NETRESOURCE(0, 0, 0, 0, None, remote, None, None)
	mpr = ctypes.windll.mpr
	code = mpr.WNetAddConnection2W(ctypes.byref(resource), password or None, username or None, 0)
	if code == 1219:
		# Already connected under a different login: drop that one and retry.
		mpr.WNetCancelConnection2W(remote, 0, True)
		code = mpr.WNetAddConnection2W(ctypes.byref(resource), password or None, username or None, 0)
	if code == 0:
		return True, "Signed in."
	if code in (1326, 86, 1331, 1330):
		return False, "The username or password was not accepted."
	if code in (53, 67, 1231):
		return False, "The router could not be reached over the network."
	return False, f"Sign-in failed (Windows error {code})."


def fingerprint_share_format(share):
	"""
	Work out a share's filesystem from how it stores timestamps -- a signal the
	router can't fake, unlike the name Windows reports (which is identical for
	every partition). Write a tiny probe file, stamp it, and see what survives:
	FAT32 stores the last-access time as a date only (time of day is lost);
	exFAT keeps time of day but only to 2 s (Windows) or 10 ms (Linux); NTFS
	keeps the full fraction. Checked against real FAT32, exFAT and NTFS
	volumes. Returns "FAT32", "exFAT", "NTFS", or "" if the result is
	inconclusive (e.g. the server ignored the timestamps).
	"""
	name = f"_fmt_probe_{os.getpid()}.tmp"
	probe = os.path.join(share, name)
	second = 1_000_000_000
	modified = 1_700_000_001_123_456_700      # odd second plus a fraction
	accessed = 1_700_000_042_000_000_000      # a time of day, well off midnight
	try:
		with open(probe, "wb") as handle:
			handle.write(b"x")
		os.utime(probe, ns=(accessed, modified))
		# Read the times back from a fresh directory listing, not a cached handle.
		found = None
		with os.scandir(share) as entries:
			for entry in entries:
				if entry.name == name:
					found = entry.stat()
					break
		if found is None or abs(found.st_mtime_ns - modified) > 3 * second:
			return ""                          # file vanished, or timestamps ignored
		if abs(found.st_atime_ns - accessed) > 3 * second:
			# Time of day lost; accept only if it fell back to within a day.
			return "FAT32" if abs(found.st_atime_ns - accessed) < 86_400 * second else ""
		if abs(found.st_mtime_ns - modified) < 1000:
			return "NTFS"
		if found.st_mtime_ns % 10_000_000 == 0:
			return "exFAT"
		return ""
	except OSError:
		return ""
	finally:
		try:
			os.remove(probe)
		except OSError:
			pass


def describe_share(share, fmt=""):
	"""
	Facts about a share for the GUI: total and free bytes, and -- when this
	PC also has the same drive plugged in -- the local drive letter that
	matches it (same filesystem and size within 2%). The router names its
	shares itself (G, H, I), so this is the only honest way to say which
	partition is which; with no local match, the keys are simply absent.
	"""
	info = {}
	try:
		usage = shutil.disk_usage(share)
	except OSError:
		return info
	info["total"], info["free"] = usage.total, usage.free
	if not fmt or platform.system() != "Windows":
		return info
	matches = []
	mask = ctypes.windll.kernel32.GetLogicalDrives()
	for index in range(26):
		if not mask & (1 << index):
			continue
		letter = chr(ord("A") + index)
		if ctypes.windll.kernel32.GetDriveTypeW(f"{letter}:\\") in (4, 5):   # network, optical
			continue
		local_fs = get_volume_filesystem_native(f"{letter}:") or ""
		if local_fs.lower() != fmt.lower():
			continue
		try:
			local_total = shutil.disk_usage(f"{letter}:\\").total
		except OSError:
			continue
		if abs(local_total - usage.total) <= 0.02 * usage.total:
			matches.append(letter)
	if len(matches) == 1:
		info["local_letter"] = matches[0]
	return info


def resolve_usb_share(config):
	"""
	Determine the USB share path for this run without requiring manual config edits.
	Works regardless of how the share was named on the router.

	Discovery order:
	  1. Configured path already accessible   -> use it silently
	  2. Scan gateway for disk shares         -> writable ones are candidates
	  3. Scan ARP-reachable subnet neighbors  -> catches Extender shares at their own IP
	  4. Multiple found -> prompt user to choose
	  5. None found    -> prompt for manual entry
	"""
	configured = config.get("usb_share_path", "")
	gateway = config.get("detected_gateway_ip", "")

	if configured and isdir_bounded(configured, timeout=8):
		return configured

	found = []
	print("Scanning network for USB shares...")

	# Scan the gateway — covers Router USB and any Extender USB proxied through the Router
	if gateway and ping_host(gateway):
		for name in _list_disk_shares(gateway):
			path = f"\\\\{gateway}\\{name}"
			if validate_share_write_access(path):
				found.append(path)

	# Scan ARP-reachable hosts in the same /24 subnet for Extender shares at their own IP
	try:
		ps = (
			"Get-NetNeighbor -State Reachable -AddressFamily IPv4 "
			"| Where-Object { $_.IPAddress -notmatch ':' } "
			"| Select-Object -ExpandProperty IPAddress"
		)
		result = subprocess.run(
			["powershell", "-NoProfile", "-Command", ps],
			capture_output=True, text=True, timeout=5, check=False,
		)
		gateway_prefix = ".".join(gateway.split(".")[:3]) + "." if gateway else ""
		neighbors = [
			ip.strip() for ip in result.stdout.splitlines()
			if ip.strip() and ip.strip() != gateway
			and (not gateway_prefix or ip.strip().startswith(gateway_prefix))
		]
		for ip in neighbors[:8]:  # cap at 8 to keep startup fast
			for name in _list_disk_shares(ip, timeout=2):
				path = f"\\\\{ip}\\{name}"
				if validate_share_write_access(path) and path not in found:
					found.append(path)
	except Exception:
		pass

	if len(found) == 1:
		print(f"USB share auto-detected: {found[0]}")
		return found[0]

	if len(found) > 1:
		print(f"Multiple USB shares detected: {', '.join(found)}")
		return prompt_choice("Which USB share to test", found, found[0])

	print(
		f"No USB share detected automatically. Check the drive is plugged into the router "
		f"and USB sharing is on; continuing with {configured or 'no share'}, which will be "
		f"reported as not accessible if it still can't be reached."
	)
	raw = prompt_value(
		"Enter USB share path (e.g. \\\\192.168.1.1\\share  or  \\\\MyRouter\\media)",
		configured,
	).strip()
	if raw and not raw.startswith("\\\\"):
		raw = normalize_share_path(raw, config.get("usb_share_name", "g"))
	return raw


def _execute_round_matrix(run_configs, interfaces, iters, total, done, started, all_results):
	"""
	Run the full target x interface x format x file x direction x iteration
	matrix for the given interfaces, appending results to all_results and
	returning the updated `done` count. Shared by both the normal
	single-connection flow and the per-band loop used when auto-switching
	wireless bands, so the actual test-execution logic exists in one place.

	If the user presses Stop, the rounds already measured are kept, and every
	interface that did not finish is recorded as SKIPPED so the report can't
	pass off a partial run as a complete one.
	"""
	progress = {"done": done, "cfg": 0, "iface": 0}
	try:
		return _run_round_matrix(run_configs, interfaces, iters, total, done, started, all_results, progress)
	except StopRequested:
		print("\nStop requested -- finishing up and saving what was measured so far.")
		logging.info("Stop requested by the user")
		reason = "Run stopped by the user before this interface finished."
		for cfg_index in range(progress["cfg"], len(run_configs)):
			first = progress["iface"] if cfg_index == progress["cfg"] else 0
			for interface in interfaces[first:]:
				all_results.append(skipped_result(run_configs[cfg_index], interface.get("name", "unknown"), reason))
		return progress["done"]


def _run_round_matrix(run_configs, interfaces, iters, total, done, started, all_results, progress):
	for cfg_index, run_cfg in enumerate(run_configs):
		progress["cfg"] = cfg_index
		target_role = run_cfg.get("dut", {}).get("role", "Unknown")
		target_share = run_cfg.get("usb_share_path", "")
		if not wait_for_share(target_share):
			hint = " Check that the drive is plugged in and shared."
			print(f"\nSkipping target {target_role}: {target_share} is not accessible.{hint}")
			print(f"  -> Update usb_share_path or dut_targets in test_config.json to match the correct address.")
			logging.warning(f"Target skipped: {target_role} share {target_share} not accessible")
			for interface in interfaces:
				all_results.append(skipped_result(
					run_cfg, interface.get("name", "unknown"),
					f"Target {target_role} not tested: share {target_share} is not accessible or not writable.{hint}",
				))
			continue
		logging.info(f"Target: {target_role} share={target_share}")
		for iface_index, interface in enumerate(interfaces):
			progress["iface"] = iface_index
			detected_formats = active_usb_formats(run_cfg)
			for usb_format in detected_formats:
				for test_file in run_cfg.get("test_files", []):
					for direction in ["read", "write"]:
						for iteration in range(1, iters + 1):
							label = f"[{target_role}] {interface.get('name')} {usb_format} {direction} {test_file['name']} round {iteration}"
							if stop_requested():
								raise StopRequested()
							show_progress(done, total, label, started)
							current_result = run_case(run_cfg, interface, usb_format, direction, test_file, iteration, label)
							all_results.append(current_result)
							done += 1
							progress["done"] = done
							save_checkpoint(done)
							logging.info(
								f"{target_role} {direction.upper()} "
								f"{test_file['name']} {usb_format} "
								f"{current_result.status} "
								f"{current_result.throughput_mbps:.2f} MB/s"
							)
							show_progress(done, total, label, started)
	return done


def run(config_path="test_config.json"):
	config_path = resolve_config_path(config_path)
	os.chdir(config_path.resolve().parent)
	setup_logging()
	logging.info("USB performance test started")
	# A stop flag older than this process is left over from an earlier run.
	# (One written after we started is a real Stop press, so keep it.)
	try:
		flag = resolve_path(STOP_FLAG)
		if os.path.exists(flag) and os.path.getmtime(flag) < _PROCESS_START:
			os.remove(flag)
	except OSError:
		pass
	print("[1/4] Checking the network and USB share...")
	config = load_config(config_path)
	global _ip_probe_target
	share_ip = share_host(config.get("usb_share_path", ""))
	if re.match(r"^\d+\.\d+\.\d+\.\d+$", share_ip):
		_ip_probe_target = share_ip
	# The slow, independent checks (each is a separate Windows tool that takes
	# a second or more) run side by side instead of one after another, and the
	# file hashing proceeds in the background while the router is being read.
	background = ThreadPoolExecutor(max_workers=4)
	connection_future = background.submit(active_connection_type)
	wifi_future = background.submit(wifi_info)
	link_future = background.submit(client_link_info, config.get("client", {}).get("ip_address") or client_ip())
	config["test_files"] = discover_test_files(config)
	verify_future = background.submit(verify_test_files, config, config["test_files"])
	gateway = gateway_ip(near=share_ip)
	config["detected_gateway_ip"] = gateway
	ping_future = background.submit(ping_host, gateway)
	# Resolve the USB share path at runtime — no manual config edit needed
	config["usb_share_path"] = resolve_usb_share(config)
	config["share_write_access"] = validate_share_write_access(
	config.get("usb_share_path", "")
	)
	config["dut_reachable"] = ping_future.result()
	print(f"[2/4] Reading router model and firmware from {config.get('detected_gateway_ip') or 'the gateway'} (a browser window may open; log in if asked)...")
	config["dut"] = detect_dut_info(config)
	logging.info(
		f"DUT info source: {config['dut'].get('info_source', 'unknown')}; "
		f"model={config['dut'].get('model')}, name={config['dut'].get('name', '')}, "
		f"firmware={config['dut'].get('firmware_version')}, hw={config['dut'].get('hardware_version', '')}"
	)
	print("[3/4] Verifying the test file set...")
	config["file_set_summary"] = verify_future.result()
	print(f"Test file set: {config['file_set_summary']}")
	logging.info(f"Test file set: {config['file_set_summary']}")
	logging.info(f"Gateway detected: {gateway}")
	interfaces = normalize_interfaces(config.get("interfaces", []))
	active_type = config.get("active_connection_type", "all").lower()
	print("[4/4] Detecting the connection type...")
	detected_type = connection_future.result()
	config["actual_connection_type"] = detected_type
	wifi = wifi_future.result()
	config["actual_wifi"] = wifi
	background.shutdown(wait=False)
	config["topology"] = derived_topology(config)
	effective_type = (
		detected_type
		if active_type == "auto" or (active_type == "all" and config.get("skip_inactive_interfaces", True))
		else active_type
	)
	interfaces = filter_interfaces(interfaces, effective_type, wifi)
	link_name, link_mbps = link_future.result()
	config["client_link"] = {"adapter": link_name, "mbps": link_mbps}
	if detected_type == "wired" and link_mbps and link_mbps < 1000:
		cap = link_mbps / 8 * 0.94
		warning = (f"This PC's Ethernet link ({link_name}) is only {link_mbps:g} Mbps, so wired speeds are capped "
		           f"near {cap:.0f} MB/s no matter what the router or the drive can do. Check the cable and the "
		           f"port, or the adapter's Speed & Duplex setting, and run again.")
		config["client_link_warning"] = warning
		print("⚠  " + warning)
		logging.warning(warning)
	elif detected_type == "wired" and link_mbps:
		# The link is gigabit-class or better but slower than the router port that was
		# selected (e.g. a 1 Gbps adapter on a 10 Gbps port): not a fault, but large
		# reads can be limited by this PC, so say so alongside the results.
		matches = (re.search(r"([\d.]+)\s*Gbps", item.get("name", ""), re.IGNORECASE)
		           for item in interfaces if item.get("connection_type") == "wired")
		ports = [float(m.group(1)) * 1000 for m in matches if m]
		if ports and link_mbps < max(ports):
			note = (f"This PC's link ({link_name}) is {link_mbps / 1000:g} Gbps while the selected router port is "
			        f"{max(ports) / 1000:g} Gbps, so large-file speeds can be limited to about "
			        f"{link_mbps / 8 * 0.94:.0f} MB/s by this PC. A faster adapter is needed to measure the router fully.")
			config["client_link_note"] = note
			print("ℹ  " + note)
			logging.info(note)
	# When wired and multiple LAN ports are configured, ask which port the cable is in.
	# Windows cannot detect the specific port number, so accurate labeling requires a prompt.
	if effective_type == "wired":
		wired_ports = [item["name"] for item in interfaces if item.get("connection_type") == "wired"]
		if len(wired_ports) > 1:
			active_port = prompt_choice(
				"Multiple LAN ports configured — which port is the Ethernet cable in",
				wired_ports, wired_ports[0]
			)
			if active_port:
				interfaces = [item for item in interfaces if item.get("connection_type") != "wired" or item["name"] == active_port]
	logging.info(f"Connection type: {detected_type}")
	logging.info(f"SSID: {wifi.get('ssid','')}")
	logging.info(f"Band: {wifi.get('band','')}")
	print(f"Active client connection: {detected_type}; running: {effective_type if effective_type != 'all' else 'all configured interfaces'}")
	# Client IP not obtained: every round would fail the same way, so record
	# it once, clearly, instead (band auto-switching gets its own IP per band).
	ip_problem = client_ip_problem(config.get("client", {}).get("ip_address") or client_ip())
	early_skips = []
	if ip_problem and not (config.get("auto_switch_wireless_bands", False) and effective_type in ("wireless", "all")):
		print(f"Client IP not obtained: {ip_problem}. Tests were not run -- reconnect and try again.")
		logging.error(f"Client IP not obtained: {ip_problem}")
		for item in interfaces:
			early_skips.append(skipped_result(config, item.get("name", "unknown"), f"Client IP not obtained: {ip_problem}."))
		interfaces = []
	# Build per-target run configs (multi-target when dut_targets is configured)
	configured_targets = [t for t in config.get("dut_targets", [])
	                       if t.get("usb_share_path") or t.get("ip_address")]
	run_configs = (
		[_build_target_config(config, t) for t in configured_targets]
		if configured_targets else [config]
	)
	if len(run_configs) > 1:
		roles_str = ", ".join(rc.get("dut", {}).get("role", "?") for rc in run_configs)
		print(f"Multi-target run: {roles_str}")
	iters = int(config.get("iterations", 3))
	formats_count = (
		len(config.get("usb_formats", []))
		if config.get("test_all_configured_usb_formats", False)
		else 1
	)
	all_results = list(early_skips)
	done = 0
	started = time.perf_counter()

	auto_switch = config.get("auto_switch_wireless_bands", False)
	band_ssids = config.get("wireless_band_ssids", {})

	if auto_switch and effective_type in ("wireless", "all"):
		# Consider the FULL configured interface list (before single-band
		# filtering) so bands other than the currently-connected one are
		# still candidates -- that's the whole point of auto-switching.
		all_configured_interfaces = normalize_interfaces(config.get("interfaces", []))
		wireless_candidates = [
			item for item in all_configured_interfaces
			if item.get("connection_type") == "wireless" and band_ssids.get(item.get("name", ""))
		]
		wired_interfaces = [item for item in interfaces if item.get("connection_type") == "wired"]
		for item in all_configured_interfaces:
			if item.get("connection_type") == "wireless" and not band_ssids.get(item.get("name", "")):
				all_results.append(skipped_result(
					config, item.get("name", "unknown"),
					f"Band not tested: no SSID configured for {item.get('name', 'this band')}.",
				))

		wireless_passwords = {}
		passwords_env = os.environ.get("USB_TEST_WIFI_PASSWORDS", "")
		if passwords_env:
			try:
				wireless_passwords = json.loads(passwords_env)
			except (json.JSONDecodeError, TypeError):
				print(
					"Warning: USB_TEST_WIFI_PASSWORDS could not be parsed; "
					"proceeding without passwords (a pre-saved profile will be required)."
				)

		if wireless_candidates:
			show_matrix_summary(config, wired_interfaces + wireless_candidates)
			original_wifi = wifi_info()
			per_band_rounds = len(run_configs) * formats_count * len(config.get("test_files", [])) * 2 * iters
			total = per_band_rounds * (len(wired_interfaces) + len(wireless_candidates))
			if total == 0:
				print("No test rounds matched the current connection/filter settings; writing empty reports.")

			if wired_interfaces:
				done = _execute_round_matrix(run_configs, wired_interfaces, iters, total, done, started, all_results)

			bands_using_temp_profile = []
			for band_interface in wireless_candidates:
				band_name = band_interface.get("name", "")
				if stop_requested():
					all_results.append(skipped_result(
						config, band_name, "Run stopped by the user before this band was tested."))
					continue
				ssid = band_ssids.get(band_name, "")
				band_password = wireless_passwords.get(band_name, "")
				print(f"\nSwitching to {band_name} ({ssid})...")
				if switch_wifi_band(ssid, band_name, password=band_password or None):
					print(f"Connected to {band_name}.")
					logging.info(f"Band switch OK: {band_name} ({ssid})")
					# Right after joining, the link is still negotiating its rate
					# and DHCP/SMB are re-establishing; measuring immediately
					# under-reports the band's real speed.
					settle = int(config.get("band_settle_seconds", 5))
					if settle > 0:
						print(f"Letting the {band_name} connection settle for {settle}s before measuring...")
						time.sleep(settle)
					if band_password:
						bands_using_temp_profile.append(ssid)
					done = _execute_round_matrix(run_configs, [band_interface], iters, total, done, started, all_results)
				else:
					landed = wifi_info().get("band", "") or "no connection"
					reason = _switch_detail or "reason unknown"
					print(
						f"Could not connect to {band_name} via SSID '{ssid}': {reason}. "
						f"Skipping this band; check that it is enabled in the router app "
						f"and within range."
					)
					logging.warning(f"Band switch FAILED, skipped: {band_name} ({ssid}); {reason}; landed on {landed}")
					all_results.append(skipped_result(
						config, band_name,
						f"Band not tested: could not connect to {band_name} via SSID '{ssid}' ({reason}; client ended on: {landed}).",
					))
					total -= per_band_rounds

			restore_band_preference()
			if original_wifi and original_wifi.get("ssid"):
				print(f"\nReconnecting to original network {original_wifi['ssid']}...")
				switch_wifi_band(original_wifi["ssid"], original_wifi.get("band", ""))

			# Networks connected to via a typed password get their temporary
			# profile removed after use, so the password doesn't linger in
			# Windows' own saved-profile store beyond this session -- matching
			# the same "never persists" principle already applied to
			# test_config.json.
			for ssid in bands_using_temp_profile:
				remove_temp_wifi_profile(ssid)
		else:
			print(
				"auto_switch_wireless_bands is enabled but no wireless_band_ssids are "
				"configured for the current interfaces -- falling back to normal "
				"single-connection testing."
			)
			show_matrix_summary(config, interfaces)
			total = len(run_configs) * len(interfaces) * formats_count * len(config.get("test_files", [])) * 2 * iters
			if total == 0:
				print("No test rounds matched the current connection/filter settings; writing empty reports.")
			done = _execute_round_matrix(run_configs, interfaces, iters, total, done, started, all_results)
	else:
		show_matrix_summary(config, interfaces)
		total = len(run_configs) * len(interfaces) * formats_count * len(config.get("test_files", [])) * 2 * iters
		if total == 0:
			print("No test rounds matched the current connection/filter settings; writing empty reports.")
		done = _execute_round_matrix(run_configs, interfaces, iters, total, done, started, all_results)

	print()
	save_outputs(config, all_results)
	try:
		os.remove(resolve_path(STOP_FLAG))      # a Stop press is used up once the run has ended
	except OSError:
		pass
	report_path = os.path.abspath(os.path.join(resolve_path(config.get("output_path", "test_results")), "report.html"))
	skipped_count = sum(1 for r in all_results if r.status == "SKIPPED")
	print(f"Completed {len(all_results) - skipped_count} test rounds ({skipped_count} skipped). Results: {os.path.abspath(config.get('output_path', 'test_results'))}")
	webbrowser.open(f"file:///{report_path}")


if __name__ == "__main__":
	run()