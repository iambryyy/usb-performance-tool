"""
Hermetic tests for the USB performance tool: no router, no network, no browser.

A folder on the local disk stands in for the router's USB share, and every function
that would talk to the network is replaced. Run from the project folder with:

    python -m unittest discover -s tests -v
"""
import builtins
import json
import os
import shutil
import sys
import tempfile
import threading
import time
import unittest
from http.server import BaseHTTPRequestHandler, HTTPServer

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
TMP = tempfile.mkdtemp(prefix="usb_tool_tests_")

import app_paths
app_paths.app_dir = lambda: TMP                      # logs, results and flags land in the temp folder
for _name in ("chart.umd.min.js", "file_set_manifest.json"):
    shutil.copy(os.path.join(ROOT, _name), TMP)

import usb_test as u                                  # noqa: E402  (must follow the app_dir patch)
import excel_report                                   # noqa: E402

u.webbrowser.open = lambda *a, **k: None              # never open a browser from a test


def tearDownModule():
    shutil.rmtree(TMP, ignore_errors=True)


def make_share(megabytes=(1, 2)):
    """A fresh local 'share' and 'local test files' folder pair, plus a config pointing at them."""
    root = tempfile.mkdtemp(dir=TMP)
    share, local = os.path.join(root, "share"), os.path.join(root, "local")
    os.makedirs(share)
    os.makedirs(local)
    files = []
    for mb in megabytes:
        name = f"f{mb}mb.bin"
        with open(os.path.join(local, name), "wb") as handle:
            block = os.urandom(1024 * 1024)
            for _ in range(mb):
                handle.write(block)
        files.append({"name": name, "size_mb": mb, "generate": False})
    config = {
        "local_test_file_dir": local, "output_path": os.path.join(root, "out"), "usb_share_path": share,
        "current_usb_format": "NTFS", "test_timeout_seconds": 20, "dut": {"role": "Router"}, "client": {},
        "thresholds_mbps": {"read": {"wired": 1}, "write": {"wired": 1}}, "test_files": files,
        # A local NVMe copy can beat the 10 Gbps "impossible speed" ceiling; raise it for the stand-in share.
        "max_plausible_mbps": 50000,
    }
    return share, local, config


WIRED = {"name": "1Gbps", "connection_type": "wired"}


class ThresholdTests(unittest.TestCase):
    CONFIG = {
        "thresholds_mbps": {"read": {"wired": 12, "wireless": 12}, "write": {"wired": 10, "wireless": 10}},
        "thresholds_by_file_mbps": {
            "read": {"wired": {"1": 21, "10": 53, "50": 57, "100": 58, "500": 70}},
            "write": {"wired": {"1": 18, "10": 38, "50": 45, "100": 48, "500": 49}},
        },
    }

    def test_exact_sizes(self):
        self.assertEqual([u.threshold(self.CONFIG, "read", "wired", mb) for mb in (1, 10, 50, 100, 500)],
                         [21, 53, 57, 58, 70])

    def test_between_and_outside_the_table(self):
        self.assertEqual(u.threshold(self.CONFIG, "read", "wired", 75), 57)     # next size down
        self.assertEqual(u.threshold(self.CONFIG, "read", "wired", 0.5), 21)    # below the smallest
        self.assertEqual(u.threshold(self.CONFIG, "read", "wired", 5000), 70)   # above the largest

    def test_falls_back_to_the_flat_value(self):
        self.assertEqual(u.threshold(self.CONFIG, "read", "wireless", 100), 12)   # no wireless table
        self.assertEqual(u.threshold(self.CONFIG, "write", "wired"), 10)          # no size given
        plain = {"thresholds_mbps": {"read": {"wired": 12}}}
        self.assertEqual(u.threshold(plain, "read", "wired", 500), 12)

    def test_a_malformed_table_does_not_crash(self):
        bad = {"thresholds_mbps": {"read": {"wired": 12}}, "thresholds_by_file_mbps": {"read": {"wired": {"x": "y"}}}}
        self.assertEqual(u.threshold(bad, "read", "wired", 10), 12)

    def test_threshold_text(self):
        labels = dict(u.threshold_lines(self.CONFIG))
        self.assertIn("21 / 53 / 57 / 58 / 70", labels["Pass thresholds (wired)"])
        self.assertIn("Read", labels["Pass thresholds (wireless)"])


class NetworkHelperTests(unittest.TestCase):
    def test_client_ip_problems(self):
        self.assertEqual(u.client_ip_problem("192.168.1.20"), "")
        for bad in ("", "169.254.3.4", "127.0.0.1"):
            self.assertNotEqual(u.client_ip_problem(bad), "", bad)

    def test_role_comes_from_the_gateway(self):
        original = u.resolve_host_ip
        u.resolve_host_ip = lambda host: host
        try:
            self.assertEqual(u.infer_dut_role("192.168.1.1", "192.168.1.1"), "Router")
            self.assertEqual(u.infer_dut_role("192.168.1.2", "192.168.1.1"), "Extender")
            self.assertEqual(u.infer_dut_role("", "192.168.1.1"), "Unknown")
        finally:
            u.resolve_host_ip = original

    def test_a_configured_role_wins(self):
        role, assumed = u.resolve_dut_role({"dut": {"role": "Extender"}, "usb_share_path": r"\\10.0.0.9\x"}, "10.0.0.1")
        self.assertEqual((role, assumed), ("Extender", False))

    def test_candidate_hosts_are_private_and_deduplicated(self):
        original = u.powershell
        u.powershell = lambda *a, **k: "192.168.1.1\n8.8.8.8\n169.254.1.1\n192.168.1.1\n"
        try:
            hosts, trusted = u.candidate_share_hosts(["192.168.1.1", "10.0.0.5", "not-an-ip"])
        finally:
            u.powershell = original
        self.assertEqual(hosts, ["192.168.1.1", "10.0.0.5"])
        self.assertEqual(trusted, set(hosts))


class UpnpTests(unittest.TestCase):
    XML = (b'<?xml version="1.0"?><root xmlns="urn:schemas-upnp-org:device-1-0"><device>'
           b"<friendlyName>Living Room Router</friendlyName><manufacturer>ExampleCo</manufacturer>"
           b"<modelName>X1</modelName><modelNumber>1.2.3</modelNumber></device></root>")

    def test_device_description_is_parsed(self):
        xml = self.XML

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *args):
                pass

            def do_GET(self):
                self.send_response(200)
                self.send_header("Content-Length", str(len(xml)))
                self.end_headers()
                self.wfile.write(xml)

        server = HTTPServer(("127.0.0.1", 0), Handler)
        threading.Thread(target=server.serve_forever, daemon=True).start()
        try:
            info = u.upnp_device_info("127.0.0.1", candidates=((server.server_address[1], "rootDesc.xml"),))
        finally:
            server.shutdown()
        self.assertEqual(info, {"name": "Living Room Router", "model": "ExampleCo X1", "firmware_version": "1.2.3"})

    def test_a_device_that_does_not_answer_gives_nothing(self):
        self.assertEqual(u.upnp_device_info("127.0.0.1", candidates=((9, "rootDesc.xml"),)), {})

    def test_unknown_stays_unknown(self):
        original = u.upnp_device_info
        u.upnp_device_info = lambda *a, **k: {}
        try:
            dut = u.detect_dut_info({"dut": {"role": "Router"}, "usb_share_path": r"\\192.168.1.1\x", "detected_gateway_ip": "192.168.1.1"})
        finally:
            u.upnp_device_info = original
        self.assertEqual((dut["model"], dut["firmware_version"], dut["role"]), ("Unknown", "Unknown", "Router"))


class FilesystemProbeTests(unittest.TestCase):
    def test_probe_names_a_format_and_cleans_up(self):
        folder = tempfile.mkdtemp(dir=TMP)
        before = os.listdir(folder)
        result = u.fingerprint_share_format(folder)
        self.assertIn(result, {"FAT32", "exFAT", "NTFS", ""})
        self.assertEqual(os.listdir(folder), before, "the probe file must be removed")

    def test_format_is_measured_once_per_share(self):
        folder = tempfile.mkdtemp(dir=TMP)
        calls = []
        original = u.fingerprint_share_format
        u.fingerprint_share_format = lambda share: (calls.append(share), "NTFS")[1]
        u._FORMAT_CACHE.clear()
        try:
            answers = {u.detect_usb_format({"usb_share_path": folder}) for _ in range(25)}
        finally:
            u.fingerprint_share_format = original
            u._FORMAT_CACHE.clear()
        self.assertEqual(answers, {"NTFS"})
        self.assertEqual(len(calls), 1)


class RoundTests(unittest.TestCase):
    def run_round(self, direction="write", mutate=None, patch=None):
        share, local, config = make_share()
        if mutate:
            mutate(share, config)
        undo = patch() if patch else None
        try:
            result = u.run_case(config, WIRED, "NTFS", direction, config["test_files"][1], 1)
        finally:
            if undo:
                undo()
        return result, share

    def test_a_normal_round_passes_and_leaves_nothing_behind(self):
        for direction in ("write", "read"):
            result, share = self.run_round(direction)
            self.assertEqual(result.status, "PASS", result.failure_reason)
            self.assertEqual(os.listdir(share), [], f"{direction} left files on the drive")

    def test_an_unreachable_share_is_recorded(self):
        result, _ = self.run_round(mutate=lambda share, cfg: cfg.update(usb_share_path=share + "_gone"))
        self.assertEqual(result.status, "ERROR")
        self.assertIn("not mounted or reachable", result.failure_reason)

    def test_a_full_drive_is_recorded(self):
        def patch():
            original = shutil.disk_usage
            shutil.disk_usage = lambda path: type("D", (), {"free": 1024})()
            return lambda: setattr(shutil, "disk_usage", original)
        result, _ = self.run_round(patch=patch)
        self.assertEqual(result.status, "ERROR")
        self.assertIn("Insufficient USB storage", result.failure_reason)

    def test_an_impossible_speed_is_not_a_pass(self):
        def patch():
            original = u.copy_timed
            u.copy_timed = lambda *a, **k: (99999.0, 0.001)
            return lambda: setattr(u, "copy_timed", original)
        result, _ = self.run_round(patch=patch)
        self.assertEqual(result.status, "ERROR")
        self.assertIn("Abnormal throughput", result.failure_reason)

    def test_a_timeout_is_recorded(self):
        def patch():
            original = u.copy_timed

            def boom(*a, **k):
                raise TimeoutError("copy stalled")
            u.copy_timed = boom
            return lambda: setattr(u, "copy_timed", original)
        result, _ = self.run_round(patch=patch)
        self.assertEqual((result.status, result.failure_reason.split(":")[0]), ("ERROR", "TimeoutError"))

    def test_below_threshold_is_a_perf_fail(self):
        result, _ = self.run_round(mutate=lambda share, cfg: cfg.update(
            thresholds_mbps={"read": {"wired": 10 ** 6}, "write": {"wired": 10 ** 6}}))
        self.assertEqual(result.status, "PERF_FAIL")

    def test_back_to_back_writes_survive_a_delete_pending_file(self):
        """Over SMB a just-deleted file can refuse to be re-created for a moment."""
        share, local, config = make_share()
        real_open, real_remove, deleted_at = builtins.open, os.remove, {}

        def slow_remove(path, *a, **k):
            real_remove(path, *a, **k)
            deleted_at[str(path)] = time.time()

        def guarded_open(path, mode="r", *a, **k):
            when = deleted_at.get(str(path))
            if when and "b" in mode and ("w" in mode or "a" in mode) and time.time() - when < 0.5:
                raise PermissionError(13, "Permission denied", path)
            return real_open(path, mode, *a, **k)

        os.remove, builtins.open = slow_remove, guarded_open
        try:
            statuses = [u.run_case(config, WIRED, "NTFS", "write", config["test_files"][0], i).status for i in (1, 2, 3)]
        finally:
            os.remove, builtins.open = real_remove, real_open
        self.assertEqual(statuses, ["PASS", "PASS", "PASS"])


class StopTests(unittest.TestCase):
    def tearDown(self):
        flag = u.resolve_path(u.STOP_FLAG)
        if os.path.exists(flag):
            os.remove(flag)

    def test_stop_keeps_finished_rounds_and_marks_the_rest_skipped(self):
        share, local, config = make_share(megabytes=(1, 24))
        interfaces = [WIRED, {"name": "2.5Gbps", "connection_type": "wired"}]
        flag, real_copy = u.resolve_path(u.STOP_FLAG), u.copy_timed

        def copy_then_press_stop(src, dst, *a, **k):
            if os.path.basename(src) == "f24mb.bin":      # Stop is pressed while the big file is moving
                open(flag, "w").close()
            return real_copy(src, dst, *a, **k)

        u.copy_timed = copy_then_press_stop
        results = []
        try:
            u._execute_round_matrix([config], interfaces, 1, 8, 0, time.perf_counter(), results)
        finally:
            u.copy_timed = real_copy
        measured = [r for r in results if r.status in ("PASS", "PERF_FAIL")]
        skipped = [r for r in results if r.status == "SKIPPED"]
        self.assertGreaterEqual(len(measured), 1)
        self.assertEqual({r.configured_interface for r in skipped}, {"1Gbps", "2.5Gbps"})
        self.assertIn("stopped by the user", skipped[0].failure_reason)
        self.assertEqual(os.listdir(share), [], "an interrupted copy must not leave files on the drive")


class EndToEndTests(unittest.TestCase):
    """The real run() twice (router, then extender) with accumulation, minus the network."""

    PATCHES = {
        "gateway_ip": lambda near="": "192.168.1.1", "ping_host": lambda ip: True,
        "active_connection_type": lambda: "wired", "client_ip": lambda: "192.168.1.50",
        "client_link_info": lambda ip: ("Ethernet", 1000.0), "upnp_device_info": lambda *a, **k: {},
        "wifi_info": lambda: {"connected": False, "ssid": "", "bssid": "", "signal": "", "band": "", "channel": ""},
    }

    def test_two_runs_merge_into_one_report(self):
        saved = {name: getattr(u, name) for name in self.PATCHES}
        for name, fake in self.PATCHES.items():
            setattr(u, name, fake)
        cwd = os.getcwd()
        share, local, base = make_share(megabytes=(1, 2))
        try:
            for role in ("Router", "Extender"):
                config = dict(base, test_files="auto", interfaces=["1Gbps"], iterations=1, accumulate_results=True,
                              active_connection_type="wired", output_path="e2e_results",
                              dut={"role": role, "name": f"{role} unit", "model": "Example", "firmware_version": "1.0"},
                              prompt_for_missing_info=False)
                path = os.path.join(TMP, "e2e_config.json")
                with open(path, "w") as handle:
                    json.dump(config, handle)
                u.run(path)
                time.sleep(1.1)                           # run files are named to the second
        finally:
            os.chdir(cwd)
            for name, original in saved.items():
                setattr(u, name, original)
        out = os.path.join(TMP, "e2e_results")
        rows = json.load(open(os.path.join(out, "test_results.json"), encoding="utf-8"))
        self.assertEqual({r["dut_role"] for r in rows}, {"Router", "Extender"})
        self.assertEqual(len(rows), 2 * 2 * 2)             # 2 roles x 2 files x read/write
        self.assertEqual(len(os.listdir(os.path.join(out, "runs"))), 2, "each run keeps its own file")
        html = open(os.path.join(out, "report.html"), encoding="utf-8").read()
        for token in ("Devices under test", "Router", "Extender", "Pass thresholds"):
            self.assertIn(token, html)
        self.assertTrue(os.path.exists(os.path.join(out, "usb_performance_results.xlsx")))
        self.assertTrue(os.path.exists(os.path.join(out, "USB_Performance_Results_TEMPLATE.xlsx")))
        self.assertEqual(os.listdir(share), [], "the drive is left clean")


class ReportTests(unittest.TestCase):
    def results(self):
        rows = []
        for role in ("Router", "Extender"):
            for fmt in ("FAT32", "exFAT"):
                for name, mb in (("small_1mb.bin", 1), ("large_500mb.bin", 500)):
                    for direction in ("read", "write"):
                        rows.append(u.Result(
                            project_name="T", dut_model="Example", dut_role=role, firmware_version="1.0", topology=f"{role} test - wired client",
                            client_type="Windows PC", client_os="Windows 11", client_ip="192.168.1.50", configured_interface="1Gbps",
                            connection_type="wired", actual_ssid="", actual_channel="", actual_band="", usb_format=fmt, direction=direction,
                            file_name=name, iteration=1, throughput_mbps=80.0 if mb > 1 else 12.0, duration_seconds=1.0, passed=True,
                            status="PASS", timestamp="2026-01-01T10:00:00", dut_name=f"{role} unit"))
        u.add_stats(rows)
        return rows

    def test_html_and_excel_are_written(self):
        out = tempfile.mkdtemp(dir=TMP)
        config = {"project_name": "T", "dut": {"role": "Router", "name": "x"}, "thresholds_mbps": {"read": {"wired": 5}, "write": {"wired": 5}}}
        u.save_html(config, self.results(), os.path.join(out, "report.html"))
        html = open(os.path.join(out, "report.html"), encoding="utf-8").read()
        self.assertIn("Devices under test", html)
        path = excel_report.write_workbook(out, self.results(), config)
        self.assertTrue(os.path.getsize(path) > 10000)


if __name__ == "__main__":
    unittest.main(verbosity=2)
