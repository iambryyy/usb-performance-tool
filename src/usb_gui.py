import json
import os
import platform
import re
import shutil
import subprocess
import sys

# When this program is relaunched as a plain process (--run-test etc.), its
# stdout is a pipe read by the GUI, and Python picks the OS's legacy codepage
# (cp1252 on Windows) for that pipe rather than UTF-8. A caught exception's
# own message can still contain characters that codepage can't represent
# (e.g. Playwright's box-drawing "install browsers" notice) -- printing that
# message to report the *original* failure would then raise a second,
# unhandled UnicodeEncodeError and crash the whole run. Swap unencodable
# characters for a placeholder instead of erroring, everywhere in this
# process, so no print() call anywhere downstream can take the program down
# over a character its console encoding doesn't happen to support.
#
# write_through: when stdout is a pipe Python block-buffers it, so any print()
# without an explicit flush (band-switch messages, share checks, router
# detection) sat in a buffer and the GUI showed nothing for a minute or more
# -- it looked frozen. Flushing every write makes output appear live.
# When the GUI launched this process it reads our output as UTF-8, so speak UTF-8
# back (symbols such as the warning sign, accented device names). The bundled
# Python ignores PYTHONIOENCODING, hence the explicit switch. A person running
# --run-test in a console keeps that console's own encoding.
_gui_pipe = os.environ.get("USB_TEST_GUI") == "1"
for _stream in (sys.stdout, sys.stderr):
    if _stream is not None:
        try:
            if _gui_pipe:
                _stream.reconfigure(encoding="utf-8", errors="replace", write_through=True)
            else:
                _stream.reconfigure(errors="replace", write_through=True)
        except Exception:
            pass

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import webbrowser

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)
from app_paths import app_dir
import usb_test
from browser_utils import launch_browser

SCRIPT_DIR = app_dir()


def downloads_folder():
    """
    This user's real Downloads folder. Asks Windows (the Known Folder API)
    rather than assuming C:\\Users\\<name>\\Downloads, which is wrong when the
    folder has been moved or redirected to another drive or a cloud-synced folder).
    """
    fallback = os.path.join(os.path.expanduser("~"), "Downloads")
    if sys.platform != "win32":
        return fallback
    try:
        import ctypes
        from ctypes import wintypes

        class GUID(ctypes.Structure):
            _fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD),
                        ("Data3", wintypes.WORD), ("Data4", ctypes.c_ubyte * 8)]

        # FOLDERID_Downloads = {374DE290-123F-4565-9164-39C4925E467B}
        folder_id = GUID(0x374DE290, 0x123F, 0x4565,
                         (ctypes.c_ubyte * 8)(0x91, 0x64, 0x39, 0xC4, 0x92, 0x5E, 0x46, 0x7B))
        path_ptr = ctypes.c_wchar_p()
        if ctypes.windll.shell32.SHGetKnownFolderPath(
                ctypes.byref(folder_id), 0, None, ctypes.byref(path_ptr)) != 0:
            return fallback
        path = path_ptr.value
        ctypes.windll.ole32.CoTaskMemFree(path_ptr)
        return path if path and os.path.isdir(path) else fallback
    except Exception:
        return fallback


def open_path(path):
    """Open a file or folder with the OS's default application."""
    if sys.platform == "win32":
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.run(["open", path], check=False)
    else:
        subprocess.run(["xdg-open", path], check=False)


def kill_process_tree(pid):
    """Force-end a process and everything it started."""
    if sys.platform == "win32":
        subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"], capture_output=True,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        return
    import signal
    try:
        os.killpg(os.getpgid(pid), signal.SIGKILL)
    except (OSError, ProcessLookupError):
        pass


def latest_results_workbook(folder):
    """
    Newest results workbook in `folder`, or None. Matches only the real results
    files (usb_performance_results.xlsx, or the time-stamped copy written when
    that file was open in Excel). A loose glob here also matched
    USB_Performance_Results_TEMPLATE.xlsx -- Windows ignores letter case -- and
    since the template is written right after the results it always looked
    newest, so "Open Excel" opened the blank template.
    """
    if not os.path.isdir(folder):
        return None
    pattern = re.compile(r"^usb_performance_results(_\d{8}_\d{6})?\.xlsx$", re.IGNORECASE)
    found = [p for p in Path(folder).iterdir() if p.is_file() and pattern.match(p.name)]
    if not found:
        return None
    return str(max(found, key=lambda p: p.stat().st_mtime))


def _relaunch_args(*flags):
    """
    Args to relaunch this same program as a plain (non-GUI) process, used to
    run usb_test.py / generate_test_files.py as a subprocess. A machine
    running the frozen .exe has no system Python to shell out to, so instead
    this always relaunches itself: sys.executable is the .exe itself when
    frozen, or the same Python interpreter otherwise, and the extra script
    path is only needed in the latter case. __main__ below checks for these
    flags and runs the requested task instead of opening the GUI.
    """
    if getattr(sys, "frozen", False):
        return [sys.executable, *flags]
    return [sys.executable, os.path.abspath(__file__), *flags]
CONFIG_FILE = os.path.join(SCRIPT_DIR, "test_config.json")
TEST_FILES_DIR = os.path.join(SCRIPT_DIR, "test_files")
REPORT_FILE = os.path.join(SCRIPT_DIR, "test_results", "report.html")

# Navy and gold palette
NAVY        = "#0B1F3A"   # header, dark surfaces
ACCENT      = "#C9A227"   # gold: primary buttons, highlights, progress
ACCENT_DARK = "#B38E1B"   # gold, pressed / hover
ACCENT_SOFT = "#E8D9A6"   # pale gold tint
WHITE       = "#FFFFFF"
BG          = "#F3F0E8"   # page background (warm parchment)
PANEL       = "#FBF9F3"   # card surface
CARD_HDR    = "#E7E1D1"   # card title bar
BORDER      = "#D9D1BD"   # borders / dividers
TEXT        = "#14233B"   # base text (navy ink)
TEXT_MUTE   = "#6E7380"   # muted / secondary text
GREEN       = "#2B8A4B"   # success indicator
ORANGE      = "#A66A1F"   # bronze: in-progress, FAT32 badge
DANGER      = "#B03A2E"   # destructive actions only


class USBPerformanceGUI:

    def __init__(self, root):
        self.root = root
        self.root.title("USB Performance Tool")
        self.root.configure(bg=BG)
        self.root.resizable(True, True)

        # Fit the window to the actual screen instead of a fixed pixel
        # size, so the bottom toolbar (Start Test, Save Config, etc.)
        # doesn't get positioned off-screen on smaller displays -- e.g. a
        # laptop's built-in screen vs. a larger external monitor.
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        win_w = min(1100, screen_w - 40)
        win_h = min(820, screen_h - 80)  # leave room for the taskbar
        self.root.geometry(f"{win_w}x{win_h}")
        self.root.minsize(700, 500)
        if platform.system() == "Windows":
            try:
                self.root.state("zoomed")  # start maximized: most reliable
                                            # fit regardless of resolution
            except tk.TclError:
                pass

        self._apply_styles()
        self._pending_wifi_passwords = {}
        self.load_config()
        self.build_ui()
        self.refresh_file_list()
        # The "This run: N rounds..." line follows every setting that changes it.
        for var in [self.iter_var, self.connection_var, self.share_var,
                    *self.band_vars.values(), *self.band_ssid_vars.values()]:
            var.trace_add("write", self._schedule_plan_refresh)
        self._refresh_result_buttons()
        self._schedule_plan_refresh()
        # Auto-populate the share dropdown on launch so the common case (drive
        # already plugged in) needs no extra clicks. The "Scan for Shares"
        # button remains available to re-scan if the drive wasn't ready yet
        # or gets swapped mid-session.
        self.scan_for_shares()

    def _apply_styles(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("Gold.Horizontal.TProgressbar", troughcolor=BORDER, background=ACCENT,
                        bordercolor=BORDER, lightcolor=ACCENT, darkcolor=ACCENT, thickness=12)
        style.configure("Treeview",
                        background=WHITE, foreground=TEXT,
                        fieldbackground=WHITE,
                        font=("Segoe UI", 9), rowheight=24)
        style.configure("Treeview.Heading",
                        background=CARD_HDR, foreground=TEXT,
                        font=("Segoe UI", 9, "bold"))
        style.map("Treeview", background=[("selected", ACCENT)],
                  foreground=[("selected", NAVY)])
        style.configure("Vertical.TScrollbar",
                        background=BORDER, troughcolor=CARD_HDR,
                        arrowcolor=TEXT_MUTE)

    def load_config(self):
        # A fresh clone ships test_config.example.json; use it until the first save
        # creates test_config.json (which stays out of version control).
        example = CONFIG_FILE.replace(".json", ".example.json")
        source = CONFIG_FILE if os.path.exists(CONFIG_FILE) else example
        if os.path.exists(source):
            with open(source, "r", encoding="utf-8") as f:
                self.config = json.load(f)
        else:
            self.config = {}

    def save_config(self):
        selected_share = self.share_var.get().strip()
        if selected_share == "All share paths" and self._discovered_shares:
            self.config["usb_share_path"] = self._discovered_shares[0]
            targets = []
            for s in self._discovered_shares:
                target = {"usb_share_path": s}
                fmt_var = self.share_format_vars.get(s)
                if fmt_var and fmt_var.get() and fmt_var.get() != "Unknown":
                    # Declared formats override auto-detection for this
                    # specific partition (see _build_target_config in
                    # usb_test.py) -- the router can't reliably report
                    # per-partition filesystem type on its own.
                    target["current_usb_format"] = fmt_var.get()
                targets.append(target)
            self.config["dut_targets"] = targets
            # Fallback for any share left "Unknown": still attempt
            # auto-detection rather than skip it entirely.
            self.config["current_usb_format"] = "auto"
        else:
            self.config["usb_share_path"] = selected_share
            self.config.pop("dut_targets", None)
            chosen_format = self.format_var.get()
            self.config["current_usb_format"] = "auto" if chosen_format == "Unknown" else chosen_format

        self.config["iterations"] = int(self.iter_var.get())
        self.config["test_timeout_seconds"] = int(self.timeout_var.get())
        self.config["accumulate_results"] = self.accumulate_var.get()

        selected = self.connection_var.get()
        self._pending_wifi_passwords = {}
        if selected == "Wireless":
            checked_bands = [b for b, v in self.band_vars.items() if v.get()]
            band_ssids = {
                b: self.band_ssid_vars[b].get().strip()
                for b in checked_bands if self.band_ssid_vars[b].get().strip()
            }
            if band_ssids:
                self.config["interfaces"] = list(band_ssids.keys())
                self.config["wireless_band_ssids"] = band_ssids
                self.config["auto_switch_wireless_bands"] = True
                # Passwords are intentionally kept out of self.config, so
                # they're never written to test_config.json. Held only for
                # this session, passed to usb_test.py as an environment
                # variable when the subprocess launches (see run_test()).
                if self.same_password_var.get():
                    shared = self.shared_password_var.get()
                    self._pending_wifi_passwords = {b: shared for b in band_ssids if shared}
                else:
                    self._pending_wifi_passwords = {
                        b: self.band_password_vars[b].get()
                        for b in band_ssids if self.band_password_vars[b].get()
                    }
            else:
                # No bands checked / no SSIDs entered -- fall back to the
                # original behavior (all bands, no auto-switching; requires
                # the client to already be connected to whichever band will
                # actually be tested).
                self.config["interfaces"] = ["2.4GHz", "5GHz", "6GHz", "MLO"]
                self.config["auto_switch_wireless_bands"] = False
        else:
            self.config["interfaces"] = [selected]
            self.config["auto_switch_wireless_bands"] = False

        self.config.setdefault("dut", {})
        self.config["dut"]["role"] = self.role_var.get() or "auto"
        for key, var in (("name", self.dut_name_var), ("model", self.dut_model_var),
                         ("firmware_version", self.dut_fw_var)):
            self.config["dut"][key] = var.get().strip() or "auto"
        # "The other device" = this PC is on the device that does not hold the drive.
        other_device = {"Router": "Extender", "Extender": "Router"}.get(self.role_var.get(), "")
        self.config["client_attached_to"] = other_device if self.attach_var.get() == "other" else ""

        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.config, f, indent=2)

        self.log("✔  Configuration saved.", tag="success")

    def scan_for_shares(self):
        threading.Thread(target=self._scan_for_shares_worker, daemon=True).start()

    def _scan_for_shares_worker(self):
        current = self.share_var.get().strip()
        host_match = re.match(r"\\\\([^\\]+)\\", current)
        # Hosts that worked before are tried first, then the default gateway, so the
        # drive is usually found with no address typed in.
        preferred = [host_match.group(1)] if host_match else []
        preferred += [h for h in self.config.get("known_share_hosts", []) if h not in preferred]
        self._ui(self.log, "Looking for the router's USB storage...", None)
        found, needs_login = usb_test.find_usb_shares(preferred)
        if not found:
            for login_host in needs_login:
                shares = self._sign_in_and_rescan(login_host)
                if shares:
                    found[login_host] = shares
                    break
        if not found:
            self._ui(self.log,
                     "\u26a0  No USB storage found. Plug the drive into the router, make sure USB sharing "
                     "is turned on in its admin page, then click Scan for Shares.", "warn")
            return
        host = host_match.group(1) if host_match and host_match.group(1) in found else next(iter(found))
        shares = found[host]
        others = [h for h in found if h != host]
        if others:
            self._ui(self.log, f"Also found USB storage on {', '.join(others)}.", None)
        ip_only = [h for h in [host] + preferred if re.match(r"^\d+\.\d+\.\d+\.\d+$", h)]
        self.config["known_share_hosts"] = list(dict.fromkeys(ip_only))[:5]
        self._discovered_shares = shares
        self._share_format_hints = self._identify_formats(shares)
        values = list(shares)
        if len(shares) > 1:
            # Multiple shares on one host usually means multiple partitions
            # of the same physical USB drive (e.g. one per filesystem).
            values = ["All share paths"] + values
        self.share_combo["values"] = values
        self.log(f"\u2714  Found {len(shares)} share(s): {', '.join(shares)}", tag="success")
        self.root.after(0, self._rebuild_share_format_panel)
        # One drive/partition found, or the one already chosen is among them:
        # fill in its role and format now so nothing has to be picked by hand.
        single = current if current in shares else (shares[0] if len(shares) == 1 else "")
        if single and self.share_var.get().strip() != single:
            self._ui(self.share_var.set, single)
        elif not single and current != "All share paths":
            # Several partitions and none chosen yet: the saved path is stale or
            # missing, so default to testing every partition found.
            self._ui(self.share_var.set, "All share paths")
        self._auto_detect(single or shares[0], with_format=bool(single))

    def _probe_selected_share(self):
        """The user picked a share from the dropdown: detect its role/format."""
        share = self.share_var.get().strip()
        if share.startswith("\\\\"):
            threading.Thread(target=self._auto_detect, args=(share, True), daemon=True).start()

    def _detect_role(self, share):
        """('Router'|'Extender'|'Unknown', assumed) for the device hosting this share:
        the Router is the router at the gateway; a share on any other address is
        an Extender. Same rule the test engine uses, so the two always agree."""
        host = usb_test.share_host(share)
        if host not in self._gateway_cache:
            self._gateway_cache[host] = usb_test.gateway_ip(near=host)
        gateway = self._gateway_cache[host]
        return usb_test.resolve_dut_role(
            {"usb_share_path": share, "dut": {"role": "auto"}, "detected_gateway_ip": gateway}, host)

    def _auto_detect(self, share, with_format):
        """Fill in the DUT role (and, for one share, its USB format) from the
        drive itself. Runs off the main thread; results are applied on it."""
        try:
            role, assumed = self._detect_role(share)
            if role in ("Router", "Extender"):
                self._ui(self._apply_detected_role, role, assumed)
                self._ui(self.log, f"DUT role: {role}" + (" (assumed)" if assumed else " (detected from the share address)"), None)
            if with_format and not self.share_var.get().strip() == "All share paths":
                fmt = self._share_format_hints.get(share) or usb_test.fingerprint_share_format(share)
                if fmt:
                    self._ui(self._apply_detected_format, fmt)
                    self._ui(self.log, f"USB format: {fmt} (detected from the drive)", None)
        except Exception:
            pass          # detection is a convenience; the user can still pick by hand

    def _apply_detected_role(self, role, assumed):
        self._detected_role = role
        self.role_var.set(role)
        self.role_note.config(text="assumed" if assumed else "\u2714 detected",
                              fg=ORANGE if assumed else GREEN)

    def _apply_detected_format(self, fmt):
        self._detected_format = fmt
        self.format_var.set(fmt)
        self.format_note.config(text="\u2714 detected")

    def _confirm_role_matches_share(self):
        """Stop a mislabelled run: a report that says Router for a drive that is
        actually in the Extender is wrong in every row. If the chosen role disagrees
        with where the share really is, say so before starting."""
        role = self.role_var.get()
        share = self.share_var.get().strip()
        if role not in ("Router", "Extender"):
            # No role yet (detection hadn't finished, or couldn't tell): work it
            # out now, and only ask the person if it truly can't be determined.
            detected, assumed = self._detect_role(share) if share.startswith("\\\\") else ("Unknown", False)
            if detected in ("Router", "Extender"):
                self._apply_detected_role(detected, assumed)
                return True
            answer = messagebox.askyesnocancel(
                "Which Device Is the USB Drive In?",
                "The app couldn't tell whether the USB drive is plugged into the "
                "Router or into an Extender.\n\n"
                "Yes = Router (the main router)\nNo = Extender (range extender)\n"
                "Cancel = don't start")
            if answer is None:
                return False
            self.role_var.set("Router" if answer else "Extender")
            return True
        if not share.startswith("\\\\"):
            return True
        detected, assumed = self._detect_role(share)
        if assumed or detected not in ("Router", "Extender") or detected == role:
            return True
        answer = messagebox.askyesnocancel(
            "Check DUT Role",
            f"DUT Role is set to {role}, but the USB share {share} is on the "
            f"{detected} (it is not at the network's gateway address).\n\n"
            f"Test it as the {detected}?\n\n"
            f"Yes = use {detected}     No = keep {role}     Cancel = don't start")
        if answer is None:
            return False
        if answer:
            self._apply_detected_role(detected, False)
        return True

    def _sign_in_and_rescan(self, host):
        """The router answers but lists no USB storage -- it is most likely
        waiting for a login (as File Explorer would ask for). Ask for it once
        here instead of making the user go and sign in by hand."""
        self._ui(self.log,
                 f"{host} is online but is not showing any USB storage yet. "
                 "It may need a login first.", "warn")
        for attempt in range(3):
            creds = self._ask_share_login(host, retry=attempt > 0)
            if creds is None:
                self._ui(self.log, "Sign-in skipped.", "warn")
                return []
            ok, message = usb_test.sign_in_to_share_host(host, *creds)
            if ok:
                self._ui(self.log, "✔  Signed in. Scanning again...", "success")
                return usb_test.discover_partition_shares(host)
            self._ui(self.log, f"⚠  {message}", "warn")
            if "not accepted" not in message:
                return []
        return []

    def _ask_share_login(self, host, retry=False):
        """Show the login box on the main thread and wait for the answer.
        Returns (username, password), or None if cancelled. The password is
        used once and never written anywhere."""
        answer = {}
        finished = threading.Event()

        def show():
            dlg = tk.Toplevel(self.root)
            dlg.title("USB Storage Login")
            dlg.configure(bg=PANEL)
            dlg.resizable(False, False)
            dlg.transient(self.root)

            def close(result):
                answer["creds"] = result
                dlg.destroy()
                finished.set()

            note = ("That login was not accepted. Try again.\n\n" if retry else "")
            tk.Label(dlg, text=note + f"The router at {host} needs a login to show its USB "
                     "storage.\nEnter the USB sharing username and password from its admin page\n"
                     "(the same ones File Explorer would ask for). The password is used\n"
                     "for this sign-in only and is never saved.",
                     bg=PANEL, fg=TEXT, font=("Segoe UI", 9), justify="left"
                     ).grid(row=0, column=0, columnspan=2, padx=16, pady=(14, 8), sticky="w")
            user_var, pass_var = tk.StringVar(), tk.StringVar()
            for row, (label, var, hide) in enumerate(
                    [("Username", user_var, ""), ("Password", pass_var, "•")], start=1):
                tk.Label(dlg, text=label, bg=PANEL, fg=TEXT, font=("Segoe UI", 10)
                         ).grid(row=row, column=0, padx=(16, 8), pady=4, sticky="w")
                entry = tk.Entry(dlg, textvariable=var, show=hide, width=28, relief="solid", bd=1,
                                 bg=WHITE, fg=TEXT, font=("Segoe UI", 10))
                entry.grid(row=row, column=1, padx=(0, 16), pady=4)
                if row == 1:
                    entry.focus_set()
                entry.bind("<Return>", lambda _e: close((user_var.get().strip(), pass_var.get())))
            buttons = tk.Frame(dlg, bg=PANEL)
            buttons.grid(row=3, column=0, columnspan=2, pady=(10, 14))
            self._btn(buttons, "Sign In", lambda: close((user_var.get().strip(), pass_var.get())),
                      style="primary").pack(side="left", padx=4)
            self._btn(buttons, "Skip", lambda: close(None)).pack(side="left", padx=4)
            dlg.protocol("WM_DELETE_WINDOW", lambda: close(None))
            dlg.grab_set()

        self._ui(show)
        finished.wait()
        return answer.get("creds")

    def _identify_formats(self, shares):
        """Probe each share for its real filesystem (see
        usb_test.fingerprint_share_format). Only trusted when the probes
        disagree with each other -- identical answers everywhere would mean
        the router's own reporting is masking the difference, so say nothing
        rather than pre-fill a wrong guess."""
        from concurrent.futures import ThreadPoolExecutor
        self._ui(self.log, "Identifying each partition's format...", None)
        with ThreadPoolExecutor(max_workers=len(shares)) as pool:
            found = dict(zip(shares, pool.map(usb_test.fingerprint_share_format, shares)))
        # Size, free space and (if the drive is also plugged into this PC) the
        # matching local drive letter, shown under each share.
        self._share_details = {}
        with ThreadPoolExecutor(max_workers=len(shares)) as pool:
            for share, info in zip(shares, pool.map(
                    lambda s: usb_test.describe_share(s, found.get(s, "")), shares)):
                self._share_details[share] = info
        if len({fmt for fmt in found.values() if fmt}) < 2:
            self._ui(self.log, "Could not tell the partitions apart automatically — "
                     "please pick each format by hand.", "warn")
            return {}
        self._ui(self.log, "Partitions identified: " + ", ".join(
            f"{share.rsplit(chr(92), 1)[-1]} = {fmt or '?'}" for share, fmt in found.items()), "success")
        return {share: fmt for share, fmt in found.items() if fmt}

    def _on_connection_type_changed(self, *args):
        self.root.after(0, self._schedule_fit)
        if self.connection_var.get() == "Wireless":
            self.wireless_panel.grid()
        else:
            self.wireless_panel.grid_remove()

    def _rebuild_share_format_panel(self):
        """
        Rebuild the per-share format declaration rows, shown only when
        "All share paths" is selected. Called whenever the share selection
        changes or a new scan finds a different set of shares.
        """
        is_all_shares = self.share_var.get().strip() == "All share paths"
        self._schedule_plan_refresh()      # the number of partitions changes the plan
        self._schedule_fit()

        # The single "USB Format" selector only applies to one share at a
        # time -- redundant and potentially confusing once per-partition
        # formats are being declared below instead.
        if hasattr(self, "usb_format_label"):
            if is_all_shares:
                self.usb_format_label.grid_remove()
                self.usb_format_row.grid_remove()
            else:
                self.usb_format_label.grid()
                self.usb_format_row.grid()

        for child in self.share_format_panel.winfo_children():
            child.destroy()
        self.share_format_vars = {}

        if self.share_var.get().strip() != "All share paths" or not self._discovered_shares:
            return

        # Pre-fill from any per-target formats already declared in a
        # previous session, so the user doesn't re-declare them every time.
        existing = {
            t.get("usb_share_path"): t.get("current_usb_format", "")
            for t in self.config.get("dut_targets", [])
        }

        tk.Label(self.share_format_panel,
                 text="Format of each partition \u2014 filled in automatically when the drive can be "
                      "identified; otherwise pick it yourself.",
                 bg=PANEL, fg=TEXT_MUTE, font=("Segoe UI", 8, "bold"), justify="left", anchor="w",
                 wraplength=640).grid(row=0, column=0, columnspan=2, sticky="w", pady=(4, 4))

        for i, share in enumerate(self._discovered_shares):
            row = 1 + i
            # A format measured from the drive just now beats whatever was
            # declared by hand in an earlier session.
            measured = self._share_format_hints.get(share, "")
            details = self._share_details.get(share, {})

            # One compact line per partition: coloured format badge, share path and
            # facts on the left; the radio buttons to override on the right.
            shell = tk.Frame(self.share_format_panel, bg=BORDER)
            shell.grid(row=row, column=0, columnspan=2, sticky="ew", pady=2)
            card = tk.Frame(shell, bg=WHITE)
            card.pack(fill="x", padx=1, pady=1)
            line = tk.Frame(card, bg=WHITE)
            line.pack(fill="x", padx=8, pady=4)

            badge_colors = {"FAT32": ORANGE, "NTFS": NAVY, "exFAT": GREEN}
            tk.Label(line, text=measured or "?", bg=badge_colors.get(measured, TEXT_MUTE),
                     fg=WHITE, font=("Segoe UI", 8, "bold"), width=7, pady=1
                     ).pack(side="left", padx=(0, 8))
            tk.Label(line, text=share, bg=WHITE, fg=TEXT, font=("Segoe UI", 9, "bold")
                     ).pack(side="left")
            facts = []
            if details.get("total"):
                gib = 1024 ** 3
                facts.append(f"{details['total'] / gib:.1f} GB, {details['free'] / gib:.1f} free")
            if details.get("local_letter"):
                facts.append(f"same drive as {details['local_letter']}:")
            if facts:
                tk.Label(line, text="  \u00b7  ".join(facts), bg=WHITE, fg=TEXT_MUTE,
                         font=("Segoe UI", 8)).pack(side="left", padx=(10, 0))

            fmt_var = tk.StringVar(value=measured or existing.get(share, "Unknown") or "Unknown")
            self.share_format_vars[share] = fmt_var
            for fmt in reversed(["FAT32", "NTFS", "exFAT", "Unknown"]):
                radio = self._radio(line, fmt, fmt_var, fmt)
                radio.config(bg=WHITE, activebackground=WHITE, font=("Segoe UI", 9))
                radio.pack(side="right", padx=(6, 0))
            if measured:
                tk.Label(line, text="\u2714 detected", bg=WHITE, fg=GREEN,
                         font=("Segoe UI", 8, "bold")).pack(side="right", padx=(0, 8))

    MIN_CONSOLE_LINES = 12          # never squeeze the console below a readable size

    def _schedule_fit(self, *_):
        job = getattr(self, "_fit_job", None)
        if job:
            self.root.after_cancel(job)
        self._fit_job = self.root.after(80, self._fit_to_window)

    def _fit_to_window(self):
        """
        Use the window well. When the content is taller than the window -- a 14-inch
        laptop -- the file list shrinks first (to 3 rows) and the console only down
        to a readable 12 lines; beyond that the page simply scrolls rather than
        squeezing the console into something too small to read. When there is room
        to spare -- a big monitor -- they grow back, the console well past its default.
        """
        self._fit_job = None
        if not hasattr(self, "_canvas") or not hasattr(self, "console") or self._canvas.winfo_height() < 200:
            return
        rows, lines = int(self.file_tree.cget("height")), int(self.console.cget("height"))
        for _ in range(40):
            self.root.update_idletasks()
            gap = self._canvas.winfo_height() - self._body.winfo_reqheight()
            if gap < 0:                                   # taller than the window: shrink
                if rows > 3:
                    rows -= 1
                    self.file_tree.config(height=rows)
                elif lines > self.MIN_CONSOLE_LINES:
                    lines -= 1
                    self.console.config(height=lines)
                else:
                    break
            elif gap > 40:                                # room to spare: grow back
                if rows < 6:
                    rows += 1
                    self.file_tree.config(height=rows)
                elif lines < 30:
                    lines += 1
                    self.console.config(height=lines)
                else:
                    break
            else:
                break

    def _update_attach_label(self, *_):
        """Name the two choices after the device that actually has the drive."""
        role = self.role_var.get()
        other = {"Router": "Extender", "Extender": "Router"}.get(role)
        if other:
            self.attach_same.config(text=f"The {role}  (where the drive is)")
            self.attach_other.config(text=f"The {other}")
        else:
            self.attach_same.config(text="Same device as the drive")
            self.attach_other.config(text="The other device")

    def _toggle_show_passwords(self):
        show = "" if self.show_passwords_var.get() else "•"
        for entry in list(self.band_password_entries.values()) + [self.shared_password_entry]:
            entry.config(show=show)

    def _toggle_shared_password(self):
        if self.same_password_var.get():
            for entry in self.band_password_entries.values():
                entry.grid_remove()
            self.shared_password_label.grid(
                row=self.shared_password_row, column=0, sticky="w", pady=(0, 4))
            self.shared_password_entry.grid(
                row=self.shared_password_row, column=1, columnspan=2, sticky="ew",
                padx=(4, 0), pady=(0, 4))
        else:
            self.shared_password_label.grid_remove()
            self.shared_password_entry.grid_remove()
            for entry in self.band_password_entries.values():
                entry.grid()

    # ── Widget helpers ────────────────────────────────────────────────────────

    def _card(self, parent, title, row, col, rowspan=1, colspan=1,
              padx=(8, 8), pady=(6, 6)):
        """White card with a thin border and a section title bar."""
        shell = tk.Frame(parent, bg=BORDER)
        shell.grid(row=row, column=col, rowspan=rowspan, columnspan=colspan,
                   sticky="nsew", padx=padx, pady=pady)
        inner = tk.Frame(shell, bg=PANEL)
        inner.pack(fill="both", expand=True, padx=1, pady=1)

        hdr = tk.Frame(inner, bg=CARD_HDR, height=30)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text=title, bg=CARD_HDR, fg=TEXT,
                 font=("Segoe UI", 9, "bold"), anchor="w"
                 ).pack(side="left", padx=10, pady=5)

        body = tk.Frame(inner, bg=PANEL)
        body.pack(fill="both", expand=True, padx=12, pady=10)
        return body

    def _btn(self, parent, text, command, style="normal", width=None):
        """Flat styled button.  style: 'primary' | 'danger' | 'normal'"""
        if style == "primary":
            bg, fg, hover = ACCENT, NAVY, ACCENT_DARK
        elif style == "danger":
            bg, fg, hover = DANGER, WHITE, "#8E2B22"
        else:
            bg, fg, hover = "#E6E0CE", TEXT, "#D6CEB4"

        kw = dict(text=text, command=command,
                  bg=bg, fg=fg, activebackground=hover, activeforeground=fg,
                  disabledforeground="#A59F8C",
                  font=("Segoe UI", 9, "bold") if style == "primary" else ("Segoe UI", 9),
                  relief="flat", cursor="hand2", padx=14, pady=6, bd=0)
        if width:
            kw["width"] = width
        btn = tk.Button(parent, **kw)
        # Remembered so set_btn_enabled() can grey the button out and restore it.
        btn.base_bg, btn.hover_bg, btn.base_fg = bg, hover, fg
        btn.bind("<Enter>", lambda _: btn.config(bg=hover) if str(btn["state"]) == "normal" else None)
        btn.bind("<Leave>", lambda _: btn.config(bg=btn.base_bg) if str(btn["state"]) == "normal" else None)
        return btn

    @staticmethod
    def set_btn_enabled(btn, enabled):
        """Enable or visibly grey out a button made by _btn()."""
        if enabled:
            btn.config(state="normal", bg=btn.base_bg, fg=btn.base_fg, cursor="hand2")
        else:
            btn.config(state="disabled", bg="#EEEADB", cursor="arrow")

    def _field_row(self, parent, label, row):
        lbl = tk.Label(parent, text=label, bg=PANEL, fg=TEXT,
                 font=("Segoe UI", 10), anchor="w")
        lbl.grid(row=row, column=0, sticky="w", pady=5)
        return lbl

    def _entry(self, parent, var, row, width=None):
        kw = dict(textvariable=var, font=("Segoe UI", 10),
                  relief="solid", bd=1, bg=WHITE, fg=TEXT,
                  insertbackground=TEXT, highlightthickness=0)
        if width:
            kw["width"] = width
        e = tk.Entry(parent, **kw)
        e.grid(row=row, column=1, sticky="ew", padx=(10, 0), pady=5)
        return e

    def _radio(self, parent, text, var, value):
        # tristatevalue: Tk draws a radio as "indeterminate" (dotted) when its variable
        # equals this, and the default is the empty string -- so with no role chosen yet
        # every role button looked half-selected. Use a value nothing ever has.
        return tk.Radiobutton(parent, text=text, variable=var, value=value,
                              tristatevalue="\x00none", bg=PANEL, fg=TEXT, selectcolor=WHITE,
                              activebackground=PANEL,
                              font=("Segoe UI", 10))

    # ── Main UI ───────────────────────────────────────────────────────────────

    def build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        # Header ──────────────────────────────────────────────────────────────
        hdr = tk.Frame(self.root, bg=NAVY, height=56)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Frame(self.root, bg=ACCENT, height=3).pack(fill="x")

        tk.Label(hdr, text="USB Performance Tool",
                 bg=NAVY, fg=WHITE, font=("Segoe UI", 13)
                 ).pack(side="left", padx=(18, 0))
        tk.Label(hdr, text="v1.0", bg="#1C355F", fg=ACCENT_SOFT,
                 font=("Segoe UI", 9), padx=8, pady=3
                 ).pack(side="right", padx=16, pady=16)
        self._btn(hdr, "?  How To", self.show_how_to,
                  style="normal").pack(side="right", padx=(0, 4), pady=10)

        # Footer -- packed BEFORE the scrollable content area below, so its
        # space is always reserved first. On smaller screens where the
        # total content is taller than the window, the content area
        # becomes scrollable instead of the footer (with Start Test)
        # silently getting clipped off-screen.
        footer = tk.Frame(self.root, bg="#EAE4D3", height=52)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)
        tk.Frame(footer, bg=ACCENT, width=4).pack(side="left", fill="y")
        self.start_button = self._btn(footer, "\u25b6  Start Test", self.start_test,
                                      style="primary")
        self.start_button.pack(side="left", padx=(14, 8), pady=10)
        self._running = False
        self._stopping = False
        self._engine_started = False
        self._link_warning = ""
        self._band_connect_failures = []
        self._progress_active = False
        self._progress_total, self._progress_done = 0, 0
        self._progress_what, self._progress_eta = "", ""
        self._plan_job = None
        self._fit_job = None
        # Only usable while a test is running; ends it cleanly and keeps the results so far.
        self.stop_button = self._btn(footer, "\u25a0  Stop", self.stop_test, style="danger")
        self.stop_button.pack(side="left", padx=(0, 8), pady=10)
        self.set_btn_enabled(self.stop_button, False)
        self._btn(footer, "Clear Runs",
                  self.clear_runs).pack(side="left", padx=(0, 8), pady=10)
        self.report_button = self._btn(footer, "Open Report", self.open_report)
        self.report_button.pack(side="left", padx=(0, 8), pady=10)
        self.excel_button = self._btn(footer, "Open Excel", self.open_excel)
        self.excel_button.pack(side="left", padx=(0, 8), pady=10)
        self._btn(footer, "Open Results Folder",
                  self.open_results_folder).pack(side="left", padx=(0, 8), pady=10)
        self.pdf_button = self._btn(footer, "\u2b07  Save as PDF", self.save_report_as_pdf)
        self.pdf_button.pack(side="left", pady=10)
        self.status_var = tk.StringVar(value="Ready")
        self.status_label = tk.Label(footer, textvariable=self.status_var,
                                     bg="#EAE4D3", fg=TEXT_MUTE,
                                     font=("Segoe UI", 9))
        self.status_label.pack(side="right", padx=16)

        # Content grid -- scrollable, so content taller than the window
        # (e.g. on a small laptop screen) can be reached by scrolling
        # instead of being cut off with no way to get to it.
        canvas_frame = tk.Frame(self.root, bg=BG)
        canvas_frame.pack(fill="both", expand=True)
        canvas = tk.Canvas(canvas_frame, bg=BG, highlightthickness=0)
        vsb = tk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        body = tk.Frame(canvas, bg=BG)
        canvas_window = canvas.create_window((0, 0), window=body, anchor="nw")
        self._canvas, self._body = canvas, body

        def _on_body_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        body.bind("<Configure>", _on_body_configure)

        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.bind("<Configure>", lambda e: self._schedule_fit(), add="+")

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        # Only scroll the canvas when the mouse is actually over it, so
        # scrolling inside the Console output or Test Files list still
        # works normally rather than being hijacked by the outer canvas.
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(1, weight=1)

        # ── Card: Configuration (top-left) ────────────────────────────────────
        cfg = self._card(body, "Configuration", row=0, col=0)
        cfg.columnconfigure(1, weight=1)

        self.share_var = tk.StringVar(
            value=self.config.get("usb_share_path", "\\\\192.168.1.1\\share"))
        self._field_row(cfg, "USB Share Path", 0)
        share_row = tk.Frame(cfg, bg=PANEL)
        share_row.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=5)
        share_row.columnconfigure(0, weight=1)
        self.share_combo = ttk.Combobox(
            share_row, textvariable=self.share_var, font=("Segoe UI", 10))
        self.share_combo.grid(row=0, column=0, sticky="ew")
        self._btn(share_row, "Scan for Shares", self.scan_for_shares,
                  style="normal").grid(row=0, column=1, padx=(6, 0))
        # Populated by scan_for_shares(); holds the actual UNC paths behind
        # the "All share paths" option, since that label itself isn't a
        # real path.
        self._discovered_shares = []
        self._share_format_hints = {}
        self._share_details = {}
        self._gateway_cache = {}
        # StringVar per discovered share, holding the user-declared format
        # for that specific partition (see _rebuild_share_format_panel).
        self.share_format_vars = {}

        self._field_row(cfg, "Device Role", 1)
        # No "Auto" choice: the role is detected from where the USB share is. A
        # blank selection (nothing detected yet) is resolved when Start is clicked.
        saved_role = self.config.get("dut", {}).get("role", "")
        self.role_var = tk.StringVar(value=saved_role if saved_role in ("Router", "Extender") else "")
        role_row = tk.Frame(cfg, bg=PANEL)
        role_row.grid(row=1, column=1, sticky="w", padx=(10, 0))
        for val in ["Router", "Extender"]:
            self._radio(role_row, val, self.role_var, val).pack(
                side="left", padx=(0, 14))
        # Shown when the role was worked out from the USB share rather than picked.
        self.role_note = tk.Label(role_row, text="", bg=PANEL, fg=GREEN,
                                  font=("Segoe UI", 8, "bold"))
        self.role_note.pack(side="left")
        self._detected_role = ""
        self.role_var.trace_add("write", lambda *a: self.role_note.config(text="")
                                if self.role_var.get() != self._detected_role else None)

        self.usb_format_label = self._field_row(cfg, "USB Format", 2)
        # "Unknown" means "let the tool detect it" and is saved as "auto". Older
        # configs may hold "Other" or "auto"; both show as Unknown.
        saved_format = self.config.get("current_usb_format", "")
        self.format_var = tk.StringVar(
            value=saved_format if saved_format in ("FAT32", "NTFS", "exFAT") else "Unknown")
        self.usb_format_row = tk.Frame(cfg, bg=PANEL)
        self.usb_format_row.grid(row=2, column=1, sticky="w", padx=(10, 0))
        for fmt in ["FAT32", "NTFS", "exFAT", "Unknown"]:
            self._radio(self.usb_format_row, fmt, self.format_var, fmt).pack(
                side="left", padx=(0, 10))
        self.format_note = tk.Label(self.usb_format_row, text="", bg=PANEL, fg=GREEN,
                                    font=("Segoe UI", 8, "bold"))
        self.format_note.pack(side="left")
        self._detected_format = ""
        self.format_var.trace_add("write", lambda *a: self.format_note.config(text="")
                                  if self.format_var.get() != self._detected_format else None)
        self.share_combo.bind("<<ComboboxSelected>>", lambda e: self._probe_selected_share())

        self._field_row(cfg, "Iterations", 3)
        self.iter_var = tk.StringVar(
            value=str(self.config.get("iterations", 1)))
        self._entry(cfg, self.iter_var, 3, width=8)

        self._field_row(cfg, "Timeout (seconds)", 4)
        self.timeout_var = tk.StringVar(
            value=str(self.config.get("test_timeout_seconds", 180)))
        self._entry(cfg, self.timeout_var, 4, width=8)

        self.accumulate_var = tk.BooleanVar(
            value=self.config.get("accumulate_results", True))
        tk.Checkbutton(cfg, text="Accumulate Results Across Runs",
                       variable=self.accumulate_var,
                       bg=PANEL, fg=TEXT, selectcolor=WHITE,
                       activebackground=PANEL, font=("Segoe UI", 10)
                       ).grid(row=5, column=1, sticky="w",
                              padx=(10, 0), pady=(6, 0))

        # Optional device details for the report. Blank means "ask the device over UPnP,
        # else show Unknown" -- never a guess.
        self._field_row(cfg, "Device (optional)", 6)
        saved_dut = self.config.get("dut", {})

        def _known(value):
            return "" if str(value or "").strip().lower() in ("", "auto", "unknown") else str(value)

        self.dut_name_var = tk.StringVar(value=_known(saved_dut.get("name")))
        self.dut_model_var = tk.StringVar(value=_known(saved_dut.get("model")))
        self.dut_fw_var = tk.StringVar(value=_known(saved_dut.get("firmware_version")))
        device_row = tk.Frame(cfg, bg=PANEL)
        device_row.grid(row=6, column=1, sticky="ew", padx=(10, 0), pady=5)
        for index, (label, var) in enumerate((("Name", self.dut_name_var), ("Model", self.dut_model_var),
                                              ("Firmware", self.dut_fw_var))):
            tk.Label(device_row, text=label, bg=PANEL, fg=TEXT_MUTE, font=("Segoe UI", 9)
                     ).grid(row=0, column=index * 2, padx=(0 if index == 0 else 10, 4))
            tk.Entry(device_row, textvariable=var, width=12, font=("Segoe UI", 10), relief="solid", bd=1,
                     bg=WHITE, fg=TEXT, insertbackground=TEXT, highlightthickness=0
                     ).grid(row=0, column=index * 2 + 1, sticky="ew")
            device_row.columnconfigure(index * 2 + 1, weight=1)

        # Per-share format panel: only relevant (and only shown) when "All
        # share paths" is selected, since the router can't reliably report
        # each partition's real filesystem itself (confirmed: three
        # independent detection methods all agreed on the same wrong
        # answer for every partition). Declaring the known format per
        # share here overrides auto-detection for that specific target.
        self.share_format_panel = tk.Frame(cfg, bg=PANEL)
        self.share_format_panel.grid(row=7, column=0, columnspan=2,
                                     sticky="ew", pady=(4, 0))
        self.share_format_panel.columnconfigure(1, weight=1)
        self.share_var.trace_add("write", lambda *a: self._rebuild_share_format_panel())

        # ── Card: Connection Type (top-right) ─────────────────────────────────
        con = self._card(body, "Connection Type", row=0, col=1)
        con.columnconfigure(0, weight=1)
        self.connection_var = tk.StringVar(value="Wireless")
        con.columnconfigure(1, weight=1)
        # Wireless on its own line, the four wired port speeds two per line.
        for index, opt in enumerate(["Wireless", "1Gbps", "2.5Gbps", "10Gbps", "SFP Port"]):
            r, c, span = (0, 0, 2) if index == 0 else ((index - 1) // 2 + 1, (index - 1) % 2, 1)
            self._radio(con, opt, self.connection_var, opt).grid(
                row=r, column=c, columnspan=span, sticky="w", pady=3, padx=(0, 12))
        top_row = 3

        # Which device this PC is plugged into / joined to. It need not be the one
        # holding the drive (e.g. drive in the Router, cable in an Extender), and the report
        # should say so exactly instead of assuming they are the same device.
        # Windows cannot tell which device a cable ends up in (a wired PC looks the
        # same behind either one), so this is a statement from the person -- but two plainly named choices, with the usual
        # one already selected, so it always shows what the test will assume.
        saved_attach = self.config.get("client_attached_to", "")
        self.attach_var = tk.StringVar(value="other" if saved_attach in ("Router", "Extender") else "same")
        attach = tk.Frame(con, bg=PANEL)
        attach.grid(row=top_row + 1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        tk.Frame(attach, bg=BORDER, height=1).pack(fill="x", pady=(0, 6))
        tk.Label(attach, text="This PC connects through", bg=PANEL, fg=TEXT_MUTE,
                 font=("Segoe UI", 8, "bold"), anchor="w").pack(anchor="w")
        self.attach_same = self._radio(attach, "", self.attach_var, "same")
        self.attach_same.pack(anchor="w", pady=1)
        self.attach_other = self._radio(attach, "", self.attach_var, "other")
        self.attach_other.pack(anchor="w", pady=1)
        self.role_var.trace_add("write", self._update_attach_label)
        self._update_attach_label()

        # Wireless band panel: only relevant (and only shown) when Wireless
        # is selected. SSIDs are saved to test_config.json like any other
        # setting; passwords are deliberately kept out of self.config and
        # never written to disk -- held only in self._pending_wifi_passwords
        # for the current session, passed to usb_test.py via an environment
        # variable when the subprocess launches (see run_test()).
        saved_ssids = self.config.get("wireless_band_ssids", {})
        self.band_vars = {}
        self.band_ssid_vars = {}
        self.band_password_vars = {}
        self.band_password_entries = {}
        self.same_password_var = tk.BooleanVar(value=False)
        self.shared_password_var = tk.StringVar(value="")

        self.wireless_panel = tk.Frame(con, bg=PANEL)
        self.wireless_panel.grid(row=top_row, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.wireless_panel.columnconfigure(1, weight=1)
        self.wireless_panel.columnconfigure(2, weight=1)

        tk.Frame(self.wireless_panel, bg=BORDER, height=1).grid(
            row=0, column=0, columnspan=3, sticky="ew", pady=(0, 6))
        for column, heading in ((0, "Bands to test"), (1, "Wi-Fi name (SSID)"), (2, "Password")):
            tk.Label(self.wireless_panel, text=heading, bg=PANEL, fg=TEXT_MUTE, anchor="w",
                     font=("Segoe UI", 8, "bold")).grid(
                row=1, column=column, sticky="w", padx=(4 if column else 0, 0), pady=(0, 2))

        # One row per band: tick box, network name and password side by side.
        band_row = 2
        for band in ["2.4GHz", "5GHz", "6GHz", "MLO"]:
            checked = band in saved_ssids and bool(saved_ssids.get(band))
            band_var = tk.BooleanVar(value=checked)
            self.band_vars[band] = band_var
            tk.Checkbutton(self.wireless_panel, text=band, variable=band_var,
                           bg=PANEL, fg=TEXT, selectcolor=WHITE,
                           activebackground=PANEL, font=("Segoe UI", 9),
                           anchor="w", width=7
                           ).grid(row=band_row, column=0, sticky="w", pady=1)

            ssid_var = tk.StringVar(value=saved_ssids.get(band, ""))
            self.band_ssid_vars[band] = ssid_var
            tk.Entry(self.wireless_panel, textvariable=ssid_var, width=10,
                     font=("Segoe UI", 9), relief="solid", bd=1,
                     bg=WHITE, fg=TEXT, insertbackground=TEXT, highlightthickness=0
                     ).grid(row=band_row, column=1, sticky="ew", padx=(4, 0), pady=1)

            pw_var = tk.StringVar(value="")
            self.band_password_vars[band] = pw_var
            pw_entry = tk.Entry(self.wireless_panel, textvariable=pw_var, width=10,
                                font=("Segoe UI", 9), relief="solid", bd=1,
                                bg=WHITE, fg=TEXT, insertbackground=TEXT,
                                highlightthickness=0, show="\u2022")
            pw_entry.grid(row=band_row, column=2, sticky="ew", padx=(4, 0), pady=1)
            self.band_password_entries[band] = pw_entry
            band_row += 1

        tk.Frame(self.wireless_panel, bg=BORDER, height=1).grid(
            row=band_row, column=0, columnspan=3, sticky="ew", pady=(4, 4))
        band_row += 1

        tk.Checkbutton(self.wireless_panel, text="Same password for all",
                       variable=self.same_password_var,
                       command=self._toggle_shared_password,
                       bg=PANEL, fg=TEXT, selectcolor=WHITE,
                       activebackground=PANEL, font=("Segoe UI", 9)
                       ).grid(row=band_row, column=0, columnspan=2, sticky="w")
        # Typing a long Wi-Fi password blind invites typos that only show up as
        # a failed band connection minutes later.
        self.show_passwords_var = tk.BooleanVar(value=False)
        tk.Checkbutton(self.wireless_panel, text="Show passwords",
                       variable=self.show_passwords_var,
                       command=self._toggle_show_passwords,
                       bg=PANEL, fg=TEXT, selectcolor=WHITE,
                       activebackground=PANEL, font=("Segoe UI", 9)
                       ).grid(row=band_row, column=2, sticky="w", padx=(4, 0))
        band_row += 1

        self.shared_password_row = band_row
        self.shared_password_label = tk.Label(
            self.wireless_panel, text="Password", bg=PANEL, fg=TEXT,
            font=("Segoe UI", 9))
        self.shared_password_entry = tk.Entry(
            self.wireless_panel, textvariable=self.shared_password_var,
            font=("Segoe UI", 9), relief="solid", bd=1, bg=WHITE, fg=TEXT,
            insertbackground=TEXT, highlightthickness=0, show="\u2022")
        # Hidden by default (per-band passwords shown instead); revealed by
        # _toggle_shared_password() when "Same password for all" is checked.

        note = tk.Label(
            self.wireless_panel,
            text="Only checked bands are tested. SSIDs are saved for next time;\n"
                 "passwords are used for this run only and never saved.",
            bg=PANEL, fg=TEXT_MUTE, font=("Segoe UI", 8), justify="left")
        note.grid(row=band_row + 1, column=0, columnspan=3, sticky="w", pady=(4, 0))

        self.connection_var.trace_add("write", self._on_connection_type_changed)
        self._on_connection_type_changed()
        self._rebuild_share_format_panel()

        # ── Card: Test Files (middle, full width) ─────────────────────────────
        files = self._card(body, "Test Files (stored on this PC)", row=1, col=0, colspan=2)
        files.columnconfigure(0, weight=1)
        files.rowconfigure(0, weight=1)

        tree_wrap = tk.Frame(files, bg=PANEL)
        tree_wrap.grid(row=0, column=0, sticky="nsew")
        tree_wrap.columnconfigure(0, weight=1)
        tree_wrap.rowconfigure(0, weight=1)

        self.file_tree = ttk.Treeview(tree_wrap,
                                      columns=("filename", "size"),
                                      show="headings", height=6)
        self.file_tree.heading("filename", text="File Name")
        self.file_tree.heading("size", text="Size")
        self.file_tree.column("filename", width=500)
        self.file_tree.column("size", width=120)
        self.file_tree.tag_configure("stripe", background="#F2EEE0")
        sb = ttk.Scrollbar(tree_wrap, orient="vertical",
                           command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=sb.set)
        self.file_tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        file_btns = tk.Frame(files, bg=PANEL)
        file_btns.grid(row=1, column=0, sticky="w", pady=(10, 0))
        self._btn(file_btns, "Generate Standard Test Files",
                  self.generate_test_files).pack(side="left", padx=(0, 6))
        self._btn(file_btns, "Add Custom File(s)",
                  self.add_custom_files).pack(side="left", padx=(0, 6))
        self._btn(file_btns, "Remove Selected",
                  self.remove_selected_file, style="danger").pack(side="left")
        self.file_note = tk.Label(files, text="", bg=PANEL, fg=TEXT_MUTE, anchor="w",
                                  justify="left", font=("Segoe UI", 9))
        self.file_note.grid(row=2, column=0, sticky="w", pady=(8, 0))

        # ── Console ───────────────────────────────────────────────────────────
        con_shell = tk.Frame(body, bg=BORDER)
        con_shell.grid(row=2, column=0, columnspan=2,
                       sticky="nsew", padx=8, pady=(0, 8))
        body.rowconfigure(2, weight=0)

        # One line above the console: the plan before a run ("30 rounds, about
        # 4 GB, roughly 12 min"), then a progress bar and time left during it.
        strip = tk.Frame(con_shell, bg=PANEL)
        strip.pack(fill="x")
        self.run_info_var = tk.StringVar(value="")
        tk.Label(strip, textvariable=self.run_info_var, bg=PANEL, fg=TEXT, anchor="w",
                 font=("Segoe UI", 9)).pack(side="left", padx=10, pady=6)
        self.progress_bar = ttk.Progressbar(strip, style="Gold.Horizontal.TProgressbar",
                                            mode="determinate", maximum=1000, length=280)

        con_hdr = tk.Frame(con_shell, bg=CARD_HDR, height=30)
        con_hdr.pack(fill="x")
        con_hdr.pack_propagate(False)
        tk.Label(con_hdr, text="Console Output", bg=CARD_HDR, fg=TEXT,
                 font=("Segoe UI", 9, "bold"), anchor="w"
                 ).pack(side="left", padx=10, pady=5)
        self._btn(con_hdr, "Clear", self._clear_console
                  ).pack(side="right", padx=8, pady=4)

        self.console = tk.Text(con_shell, height=self.MIN_CONSOLE_LINES,
                               bg="#0E1A2E", fg="#E4E8F1",
                               font=("Consolas", 9),
                               insertbackground=WHITE,
                               relief="flat", bd=0, state="normal")
        self.console.pack(fill="both", expand=True, padx=1, pady=(0, 1))
        self.console.tag_config("success", foreground="#7FD1A0")
        self.console.tag_config("error",   foreground="#FF8A80")
        self.console.tag_config("warn",    foreground="#E8C766")
        self.console.tag_config("info",    foreground="#8DB4F0")

    # ── Actions ───────────────────────────────────────────────────────────────

    def _clear_console(self):
        self.console.delete("1.0", tk.END)

    def refresh_file_list(self):
        self.file_tree.delete(*self.file_tree.get_children())
        Path(TEST_FILES_DIR).mkdir(exist_ok=True)
        count, total_mb = 0, 0.0
        for i, f in enumerate(sorted(os.listdir(TEST_FILES_DIR))):
            path = os.path.join(TEST_FILES_DIR, f)
            if os.path.isfile(path):
                size_mb = os.path.getsize(path) / (1024 * 1024)
                count, total_mb = count + 1, total_mb + size_mb
                self.file_tree.insert("", "end",
                                      values=(f, f"{size_mb:.2f} MB"),
                                      tags=("stripe",) if i % 2 else ())
        # The USB drive does not need these files on it: they live here, and the
        # test copies them to the drive (write) and back (read).
        if hasattr(self, "file_note"):
            self.file_note.config(
                text=(f"{count} file(s), {total_mb:,.0f} MB on this PC.  During a test they are copied "
                      "to the USB drive and back to measure speed, then the tool's copies are deleted "
                      "from the drive." if count else
                      "No test files yet — click \"Generate Standard Test Files\"."))
        if hasattr(self, "run_info_var"):
            self._schedule_plan_refresh()

    def generate_test_files(self):
        subprocess.run(
            _relaunch_args("--generate-test-files"),
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        self.refresh_file_list()

    def add_custom_files(self):
        files = filedialog.askopenfilenames()
        Path(TEST_FILES_DIR).mkdir(exist_ok=True)
        for f in files:
            shutil.copy2(f, os.path.join(TEST_FILES_DIR, os.path.basename(f)))
        self.refresh_file_list()

    def remove_selected_file(self):
        selected = self.file_tree.selection()
        if not selected:
            return
        filename = self.file_tree.item(selected[0])["values"][0]
        path = os.path.join(TEST_FILES_DIR, filename)
        if os.path.exists(path):
            os.remove(path)
        self.refresh_file_list()

    def log(self, text, tag=None):
        self.console.insert(tk.END, text + "\n", tag or "")
        self.console.see(tk.END)
        if tag == "error":
            self._set_status("⚠  Error", DANGER)
        elif tag == "success":
            self._set_status("✔  Done", GREEN)

    def _set_status(self, text, color=None):
        self.status_var.set(text)
        if color:
            self.status_label.config(fg=color)

    def _build_subprocess_env(self):
        """
        Build the environment for the usb_test.py subprocess. WiFi
        passwords are passed this way -- in process memory only, for this
        run -- rather than through test_config.json, so they're never
        written to any file this program controls.
        """
        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"   # pipes default to the Windows code page, which turns symbols and
                                            # accented device names into "?" -- use UTF-8 on both ends
        env["USB_TEST_GUI"] = "1"       # the test may ask this window a question instead of a console
        if self._pending_wifi_passwords:
            env["USB_TEST_WIFI_PASSWORDS"] = json.dumps(self._pending_wifi_passwords)
        return env

    def _ui(self, fn, *args):
        """Run fn(*args) on the Tk main thread. Widgets are only safe to
        touch from there; queued calls run in the order they were posted."""
        self.root.after(0, lambda: fn(*args))

    def _log_reader_error(self):
        """Record an unexpected error from the output-reading thread instead
        of letting the thread die silently (which closes the pipe and then
        crashes the test process on its next print)."""
        import traceback
        text = traceback.format_exc()
        try:
            logs = os.path.join(SCRIPT_DIR, "logs")
            os.makedirs(logs, exist_ok=True)
            with open(os.path.join(logs, "gui_reader_error.txt"), "a", encoding="utf-8") as f:
                f.write(text + "\n")
        except OSError:
            pass
        self._ui(self.log, "GUI display error (details in logs\\gui_reader_error.txt); the test is still running.", "warn")

    def run_test(self):
        # Whatever happens, hand the Start button back when the run is over.
        try:
            self._run_test()
        finally:
            self._ui(self._set_running, False)

    def _set_running(self, running):
        """A second click on Start during a run would launch a second test on
        the same drive at the same time; lock the button until it finishes."""
        self._running = running
        if running:
            self._stopping = False
            self.start_button.config(state="disabled", text="Running…", bg="#CFC7AE",
                                     activebackground="#CFC7AE", cursor="arrow")
            self.start_button.unbind("<Enter>")
            self.start_button.unbind("<Leave>")
            self.set_btn_enabled(self.stop_button, True)
        else:
            self.start_button.config(state="normal", text="▶  Start Test", bg=ACCENT,
                                     activebackground=ACCENT_DARK, cursor="hand2")
            self.start_button.bind("<Enter>", lambda _: self.start_button.config(bg=ACCENT_DARK))
            self.start_button.bind("<Leave>", lambda _: self.start_button.config(bg=ACCENT))
            self.set_btn_enabled(self.stop_button, False)

    def stop_test(self):
        """Ask the running test to stop. It finishes cleanly -- removing its
        temporary files from the drive and writing the report for whatever was
        measured -- rather than being killed with nothing saved."""
        if not self._running or self._stopping:
            return
        if not messagebox.askyesno(
                "Stop Test",
                "Stop the test now?\n\nRounds already measured are kept and the report "
                "is still written. Anything unfinished is marked as skipped."):
            return
        self._stopping = True
        self.set_btn_enabled(self.stop_button, False)
        self._set_status("⏳  Stopping…", ORANGE)
        self.log("Stop requested -- finishing the current step and saving the results so far…", "warn")
        try:
            open(usb_test.resolve_path(usb_test.STOP_FLAG), "w").close()
        except OSError:
            pass
        # Safety net: a step that can't notice the request (a Wi-Fi band switch,
        # waiting on the router login window) must not leave the app stuck.
        self.root.after(120000, self._force_stop_if_needed)

    def _force_stop_if_needed(self):
        proc = getattr(self, "_proc", None)
        if self._running and self._stopping and proc is not None and proc.poll() is None:
            self.log("The test did not stop in time, so it is being ended now. "
                     "Results from this run could not be saved.", "error")
            kill_process_tree(proc.pid)

    # ── Plan summary and progress ─────────────────────────────────────────────

    _BAR_RE = re.compile(r"^\[[#\-]+\]\s+(\d+)/(\d+)\s+(.*?)(?:\s+ETA\s+([\d.]+)\s+min)?$")
    _PCT_RE = re.compile(r"^(.*?)\s+(\d+(?:\.\d+)?)%$")

    def _begin_progress(self):
        self._progress_total, self._progress_done = 0, 0
        self._progress_what, self._progress_eta = "", ""
        self.progress_bar["value"] = 0
        self.progress_bar.pack(side="right", padx=10, pady=6)
        self.run_info_var.set("Starting…")

    def _parse_progress(self, text):
        """Turn the engine's console progress lines into the progress bar:
        '[####---] 12/40 <what> ETA 9.0 min' and the per-file 'NN%' lines."""
        step = re.match(r"^\[(\d)/4\]\s+(.*)$", text)
        if step and not self._progress_total:
            self.run_info_var.set(f"Getting ready (step {step.group(1)} of 4): {step.group(2)[:80]}")
            return
        bar = self._BAR_RE.match(text)
        if bar:
            done, total = int(bar.group(1)), int(bar.group(2))
            self._progress_done, self._progress_total = done, total
            self._progress_what = bar.group(3).strip()
            eta = bar.group(4)
            self._progress_eta = f"  ·  about {float(eta):.0f} min left" if eta and float(eta) >= 1 else ""
            self._draw_progress(0.0)
            return
        pct = self._PCT_RE.match(text)
        if pct and getattr(self, "_progress_total", 0):
            self._draw_progress(float(pct.group(2)))

    def _draw_progress(self, file_percent):
        total = max(self._progress_total, 1)
        done = self._progress_done
        self.progress_bar["value"] = min(1000, (done + file_percent / 100.0) / total * 1000)
        self.run_info_var.set(
            f"{done} of {self._progress_total} rounds done{self._progress_eta}  ·  {self._progress_what[:70]}")

    def _finish_progress(self):
        total = getattr(self, "_progress_total", 0)
        done = getattr(self, "_progress_done", 0)
        if total:
            self.progress_bar["value"] = 1000 if (done >= total and not self._stopping) else done / total * 1000
            outcome = "Stopped after" if self._stopping else "Finished"
            self.run_info_var.set(f"{outcome}: {done} of {total} rounds measured")
        else:
            self.progress_bar.pack_forget()
            self.run_info_var.set("")
            self._schedule_plan_refresh()

    def _schedule_plan_refresh(self, *_):
        job = getattr(self, "_plan_job", None)
        if job:
            self.root.after_cancel(job)
        self._plan_job = self.root.after(150, self._refresh_plan)

    def _history_speed(self, kind):
        """Typical MB/s from the last run's results (harmonic mean, the right
        average for estimating time), preferring rows of the same connection
        kind; None if there is no history yet."""
        import csv
        path = os.path.join(usb_test.resolve_path(self.config.get("output_path", "test_results")), "test_results.csv")
        try:
            with open(path, newline="", encoding="utf-8") as handle:
                rows = [r for r in csv.DictReader(handle) if r.get("status") in ("PASS", "PERF_FAIL")]
        except (OSError, csv.Error):
            return None
        speeds = {"all": [], kind: []}
        for row in rows:
            try:
                value = float(row.get("throughput_mbps") or 0)
            except ValueError:
                continue
            if value > 0:
                speeds["all"].append(value)
                if row.get("connection_type") == kind:
                    speeds[kind].append(value)
        chosen = speeds[kind] if len(speeds[kind]) >= 3 else speeds["all"]
        return len(chosen) / sum(1 / v for v in chosen) if chosen else None

    def _refresh_plan(self):
        """One line above the console: how big the run about to start is."""
        self._plan_job = None
        if self._running:
            return
        self._progress_total = 0                       # a finished run's summary gives way to the plan
        self.progress_bar.pack_forget()
        try:
            iterations = int(self.iter_var.get())
            sizes = [os.path.getsize(os.path.join(TEST_FILES_DIR, name)) / (1024 * 1024)
                     for name in os.listdir(TEST_FILES_DIR)
                     if os.path.isfile(os.path.join(TEST_FILES_DIR, name))]
        except (ValueError, OSError):
            self.run_info_var.set("")
            return
        if not sizes or iterations < 1:
            self.run_info_var.set("Add or generate test files to plan a run.")
            return
        wireless = self.connection_var.get() == "Wireless"
        interfaces = max(1, sum(1 for band, var in self.band_vars.items()
                                if var.get() and self.band_ssid_vars[band].get().strip())) if wireless else 1
        targets = len(self._discovered_shares) if (
            self.share_var.get().strip() == "All share paths" and self._discovered_shares) else 1
        rounds = targets * interfaces * len(sizes) * 2 * iterations
        # A write moves the file once; a read puts a copy on the drive first and
        # then copies it back, so the pair moves it three times.
        megabytes = targets * interfaces * iterations * sum(sizes) * 3
        text = f"This run: {rounds} rounds  ·  about {megabytes / 1024:.1f} GB transferred"
        speed = self._history_speed("wireless" if wireless else "wired")
        if speed:
            minutes = (megabytes / speed + rounds) / 60
            text += "  ·  roughly " + ("under a minute" if minutes < 1 else
                    f"{minutes:.0f} min" if minutes < 90 else f"{minutes / 60:.1f} hours") + " (based on your last run)"
        self.run_info_var.set(text)

    def _refresh_result_buttons(self):
        """Report, Excel and PDF buttons are greyed out until there is
        something for them to open."""
        folder = usb_test.resolve_path(self.config.get("output_path", "test_results"))
        has_report = os.path.exists(REPORT_FILE)
        self.set_btn_enabled(self.report_button, has_report)
        self.set_btn_enabled(self.pdf_button, has_report)
        self.set_btn_enabled(self.excel_button, bool(latest_results_workbook(folder)))

    def _run_test(self):
        self._ui(self._set_status, "⏳  Starting…", ORANGE)
        self._ui(self.log, "─── Test started " + "─" * 40, "info")
        # The test engine needs a few seconds to launch before it can print
        # anything; say so now so the app never looks frozen.
        self._ui(self.log, "Starting up -- the first status line appears in a few seconds…", "info")
        self._ui(self._begin_progress)
        try:
            os.remove(usb_test.resolve_path(usb_test.STOP_FLAG))     # left over from an earlier run
        except OSError:
            pass

        proc = subprocess.Popen(
            _relaunch_args("--run-test"),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            encoding="utf-8", errors="replace",
            env=self._build_subprocess_env(),
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            start_new_session=sys.platform != "win32",   # lets kill_process_tree end the whole group
        )

        # usb_test.py's per-round progress bar writes \r (carriage return,
        # no newline) to update a single line in place in a real terminal.
        # `for line in proc.stdout` only yields on \n, so those updates were
        # invisible here for the entire run -- the GUI looked frozen even
        # while the test was actively progressing. Read char-by-char and
        # treat \r-terminated segments as in-place progress updates instead.
        self._proc = proc
        self._progress_active = False
        self._engine_started = False
        self._link_warning = ""
        self._band_connect_failures = []
        buf = ""
        # This thread only reads the pipe and hands each line to the main
        # thread; it never touches widgets itself, and one bad line can't
        # end the loop (ending it would close the pipe on the running test).
        while True:
            try:
                ch = proc.stdout.read(1)
                if ch == "":
                    if proc.poll() is not None:
                        break
                    continue
                if ch in ("\r", "\n"):
                    if buf:
                        self._ui(self._handle_output, buf, ch == "\r")
                    buf = ""
                else:
                    buf += ch
            except Exception:
                self._log_reader_error()
                buf = ""
        if buf:
            self._ui(self._handle_output, buf, False)
        proc.wait()
        self._ui(self._finish_run)

    def _finish_run(self):
        # Runs on the main thread after every queued output line, so the
        # collected band failures are complete by now.
        self.log("─── Test finished " + "─" * 39, tag="info")
        if self._stopping:
            self._set_status("■  Stopped — results so far saved", ORANGE)
        else:
            self._set_status("✔  Test complete", GREEN)
        self._finish_progress()
        self._refresh_result_buttons()

        # A slow PC network link limits every wired result; don't let that be
        # buried in the console.
        if self._link_warning:
            messagebox.showwarning(
                "Slow Network Link",
                self._link_warning + "\n\nThe wired results from this run mostly measure "
                "that link, not the router or the drive. The report says so at the top.")

        # Band-connection failures (wrong password, band unavailable, etc.)
        # already print to the console -- pop a summary too so they aren't
        # missed among everything else that scrolled by during the run.
        if self._band_connect_failures:
            failures = list(self._band_connect_failures)
            messagebox.showwarning(
                "Wireless Band Connection Issues",
                "Some wireless bands failed to connect during this run:\n\n"
                + "\n\n".join(failures) +
                "\n\nThat band was skipped; other bands still ran normally."
            )

    def _handle_output(self, text, is_progress):
        stripped = text.rstrip()
        if not stripped:
            return
        if not self._engine_started:
            self._engine_started = True
            self._set_status("⏳  Stopping…" if self._stopping else "⏳  Running…", ORANGE)
        self._parse_progress(stripped)
        if is_progress:
            self.log_progress(stripped)
            return
        self._progress_active = False
        if stripped.startswith("Could not connect to"):
            self._band_connect_failures.append(stripped)
        if "Ethernet link" in stripped and "capped" in stripped:
            self._link_warning = stripped.lstrip("⚠ ").strip()
        lower = stripped.lower()
        if stripped.startswith("⚠"):
            tag = "warn"
        elif any(w in lower for w in ["error", "failed", "exception", "traceback"]):
            tag = "error"
        elif any(w in lower for w in ["pass", "success", "complete", "✔"]):
            tag = "success"
        elif any(w in lower for w in ["warn", "skip", "missing", "not found"]):
            tag = "warn"
        else:
            tag = None
        self.log(stripped, tag=tag)

    def log_progress(self, text):
        """Update the console's last line in place (used for \\r-based
        progress bar output) instead of appending a new line every update."""
        if self._progress_active:
            self.console.delete(self._progress_line_index, tk.END)
        self._progress_line_index = self.console.index(tk.END)
        self.console.insert(tk.END, text + "\n")
        self.console.see(tk.END)
        self._progress_active = True

    def start_test(self):
        if self._running:
            return
        if not self._confirm_wireless_passwords():
            return
        if not self._confirm_role_matches_share():
            return
        # Read the GUI's settings here on the main thread; Tk variables and
        # widgets must not be touched from the background worker thread.
        self.save_config()
        self._set_running(True)
        threading.Thread(target=self.run_test, daemon=True).start()
        # On a short screen the page scrolls; bring the progress line and console into
        # view so a started test is visibly doing something.
        self.root.after(200, lambda: self._canvas.yview_moveto(1.0))

    def _confirm_wireless_passwords(self):
        """A checked band with no password only works if a WiFi profile for
        that SSID is already saved on this PC -- warn up front instead of
        letting the user find out from console output deep into the run."""
        if self.connection_var.get() != "Wireless":
            return True
        missing = []
        for band, band_var in self.band_vars.items():
            if not band_var.get():
                continue
            ssid = self.band_ssid_vars[band].get().strip()
            if not ssid:
                continue
            if self.same_password_var.get():
                has_password = bool(self.shared_password_var.get())
            else:
                has_password = bool(self.band_password_vars[band].get())
            if not has_password:
                missing.append(f"{band} ({ssid})")
        if not missing:
            return True
        return messagebox.askyesno(
            "No Password Entered",
            "No password was entered for:\n\n  " + "\n  ".join(missing) +
            "\n\nThis only works if a WiFi profile for that SSID is "
            "already saved on this PC.\n\nContinue anyway?")

    def open_excel(self):
        folder = usb_test.resolve_path(self.config.get("output_path", "test_results"))
        workbook = latest_results_workbook(folder)
        if not workbook:
            messagebox.showwarning("No Excel Workbook", "No Excel workbook found yet.\nRun a test first.")
            return
        open_path(workbook)

    def open_report(self):
        report = os.path.abspath(REPORT_FILE)
        if os.path.exists(report):
            webbrowser.open(Path(report).as_uri())
        else:
            messagebox.showwarning(
                "No Report",
                "No report.html found yet.\nRun a test first.")

    def open_results_folder(self):
        # The folder only exists once a test has written into it (and "Clear
        # Runs" deletes it), so create it rather than showing an error.
        path = usb_test.resolve_path(self.config.get("output_path", "test_results"))
        os.makedirs(path, exist_ok=True)
        open_path(path)

    def clear_runs(self):
        """Wipe test_results/ (what accumulate_results merges across) and
        logs/ (per-run diagnostic trail) so the next run starts completely
        clean instead of mixing in data from prior runs."""
        results_dir = os.path.join(SCRIPT_DIR, "test_results")
        logs_dir = os.path.join(SCRIPT_DIR, "logs")
        if not os.path.exists(results_dir) and not os.path.exists(logs_dir):
            messagebox.showinfo("Clear Runs", "There are no previous runs to clear.")
            return
        if not messagebox.askyesno(
                "Clear Runs",
                "This will permanently delete all previous test runs, "
                "reports, and logs.\n\nContinue?"):
            return
        shutil.rmtree(results_dir, ignore_errors=True)
        shutil.rmtree(logs_dir, ignore_errors=True)
        self.log("─── Previous runs and logs cleared ───", tag="info")
        self._set_status("Runs cleared", TEXT_MUTE)
        self._refresh_result_buttons()
        self._schedule_plan_refresh()

    def show_how_to(self):
        messagebox.showinfo(
            "How To Use This Tool",
            "HOW TO RUN A TEST\n\n"
            "1. Plug the USB drive into the router. Pick its share in\n"
            "   \"USB Share Path\" (found automatically).\n\n"
            "2. Choose how this PC is connected:\n"
            "   • Wired: pick the port speed you plugged into.\n"
            "   • Wireless: tick each band to test, then enter that\n"
            "     Wi-Fi name (SSID) and password.\n\n"
            "3. Click ▶ Start Test.\n\n"
            "4. Optional: type the device's name, model and firmware so they\n"
            "   appear in the report. If you leave them blank the tool asks\n"
            "   the device over UPnP, or shows \"Unknown\" -- it never guesses.\n\n"
            "5. Wait for \"Test complete\", then click Open Report\n"
            "   (charts) or Open Excel (summary dashboard + data).\n\n"
            "TIPS\n"
            "• The Test Files list is the files on THIS PC. You don't put\n"
            "  them on the USB drive; the test copies them there and back\n"
            "  to time it, then deletes its copies from the drive.\n"
            "• Device Role and USB Format fill in automatically from the\n"
            "  drive (\"detected\"); you can still change them.\n"
            "• Empty file list? Click \"Generate Standard Test Files\".\n"
            "• Left a Wi-Fi password blank? That only works if this PC\n"
            "  has already saved that Wi-Fi network.\n"
            "• A band that can't connect is skipped; you get a summary\n"
            "  when the run finishes.\n"
            "• \"Stop\" ends a test early. Results measured so far are\n"
            "  saved and the rest is marked skipped.\n"
            "• \"Show passwords\" (Wireless) lets you check what you typed.\n"
            "• \"Clear Runs\" deletes old results, reports, and logs.\n\n"
            "More detail: docs/USER_GUIDE.md"
        )

    def save_report_as_pdf(self):
        report = os.path.abspath(REPORT_FILE)
        if not os.path.exists(report):
            messagebox.showwarning(
                "No Report",
                "No report.html found yet.\nRun a test first.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialdir=downloads_folder(),
            initialfile="usb_performance_report.pdf",
            title="Save Report as PDF",
        )
        if not save_path:
            return

        self.log("⏳  Generating PDF — please wait...", tag="info")
        self._set_status("⏳  Generating PDF…", ORANGE)

        def _generate():
            try:
                from playwright.sync_api import sync_playwright
                with sync_playwright() as p:
                    browser = launch_browser(p, headless=True)
                    # Wide viewport ensures charts render at full size before PDF export
                    page = browser.new_page(viewport={"width": 1280, "height": 900})
                    page.goto(f"file:///{report}")
                    # Wait for Chart.js to finish rendering all canvases
                    page.wait_for_load_state("networkidle")
                    page.wait_for_timeout(1500)
                    page.pdf(
                        path=save_path,
                        format="A3",
                        landscape=True,
                        print_background=True,
                        margin={"top": "12mm", "bottom": "12mm",
                                "left": "10mm", "right": "10mm"},
                    )
                    browser.close()
                self.root.after(0, lambda: self.log(f"✔  PDF saved: {save_path}", tag="success"))
                self.root.after(0, lambda: self._set_status("✔  PDF saved", GREEN))
            except Exception as error:
                # Capture the text now: Python deletes `error` when this block ends, so a
                # delayed callback that referred to it crashed and the failure was never shown.
                reason = f"{type(error).__name__}: {error}".splitlines()[0]
                self.root.after(0, lambda: self.log(f"PDF generation failed: {reason}", tag="error"))
                self.root.after(0, lambda: self._set_status("⚠  PDF failed", DANGER))

        threading.Thread(target=_generate, daemon=True).start()


if __name__ == "__main__":
    # This same program relaunches itself headlessly for a test run (see
    # _relaunch_args above) since a machine running the frozen .exe has no
    # system Python to run usb_test.py as a separate script. Handle that
    # before touching Tk so this path never opens a window.
    if "--run-test" in sys.argv:
        class _TolerantStream:
            """Progress text is only a courtesy to whoever is watching. If
            the reader on the other end of the pipe goes away, writing
            fails with an OSError -- swallow that so the test keeps running
            and still writes its results and report."""

            def __init__(self, stream):
                self._stream = stream

            def write(self, text):
                try:
                    return self._stream.write(text)
                except (OSError, ValueError):
                    return len(text)

            def flush(self):
                try:
                    self._stream.flush()
                except (OSError, ValueError):
                    pass

            def __getattr__(self, name):
                return getattr(self._stream, name)

        if sys.stdout is not None:
            sys.stdout = _TolerantStream(sys.stdout)
        usb_test.run()
    elif "--generate-test-files" in sys.argv:
        import generate_test_files
        generate_test_files.main()
    elif "--selftest-paths" in sys.argv:
        # Verifies the packaging-sensitive path lookups (chart.js, the file
        # set manifest, app_dir itself) without touching the network/browser
        # step, so this can be checked quickly and deterministically.
        tag = usb_test._chart_js_tag()
        manifest = usb_test.load_file_manifest()
        print("app_dir:", app_dir())
        print("frozen:", getattr(sys, "frozen", False))
        print("chart.js embedded (not CDN fallback):", "cdn.jsdelivr.net" not in tag, "len:", len(tag))
        print("manifest loaded:", manifest is not None, "version:", (manifest or {}).get("version"))
    elif "--selftest-browser" in sys.argv:
        # Headless (no window) check that the browser used for router
        # detection and PDF export can actually be launched and can print a PDF.
        from playwright.sync_api import sync_playwright
        import tempfile
        with sync_playwright() as p:
            browser = launch_browser(p, headless=True)
            page = browser.new_page()
            page.set_content("<h1>browser self-test</h1>")
            pdf_path = os.path.join(tempfile.gettempdir(), "usb_tool_browser_selftest.pdf")
            page.pdf(path=pdf_path)
            browser.close()
        print("browser launch ok; pdf bytes:", os.path.getsize(pdf_path))
        os.remove(pdf_path)
    elif "--selftest-excel" in sys.argv:
        # Confirms the Excel writer is bundled and works: builds a one-row workbook and the blank template.
        import tempfile
        import excel_report
        folder = tempfile.mkdtemp(prefix="usb_tool_excel_selftest_")
        row = {"dut_role": "Router", "configured_interface": "5GHz", "connection_type": "wireless", "usb_format": "NTFS",
               "direction": "read", "file_name": "small_1mb.bin", "iteration": 1, "throughput_mbps": 20.0,
               "duration_seconds": 0.05, "status": "PASS", "timestamp": "2026-01-01T00:00:00"}
        made = excel_report.write_workbook(folder, [row], {}), excel_report.write_template(folder)
        print("excel ok; files:", ", ".join(f"{os.path.basename(p)} ({os.path.getsize(p)} bytes)" for p in made))
        shutil.rmtree(folder, ignore_errors=True)
    else:
        root = tk.Tk()
        app = USBPerformanceGUI(root)
        root.mainloop()