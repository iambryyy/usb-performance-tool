# Configuration reference

The app reads `test_config.json` next to the program (a fresh clone starts from
`test_config.example.json`) and rewrites it when you press **Start Test**, so the GUI and the file
never disagree. The command-line engine (`python usb_test.py`) reads the same file.
Passwords are never stored in it.

## Where and what

| Key | Meaning |
|---|---|
| `usb_share_path` | UNC path of the share, e.g. `\\192.168.1.1\disk1`. |
| `usb_share_name` | Share name used when only an address is given. |
| `dut_targets` | Optional list of extra shares tested in the same run (one per partition or device). Each entry may set `usb_share_path`, `current_usb_format`, `role`, `name`, `model`, `firmware_version`, `ip_address`. The GUI fills this in for *All share paths*. |
| `known_share_hosts` | Addresses that worked before; tried first when scanning. Maintained by the app. |
| `output_path` | Where results go (default `test_results`). |
| `local_test_file_dir` | Folder of local test files (default `test_files`). |
| `test_files` | `"auto"` (every file in the folder) or an explicit list of `{name, size_mb, generate}`. |
| `test_file_set_version` | Version label checked against `file_set_manifest.json`. |

## The device

```json
"dut": { "role": "auto", "name": "auto", "model": "auto", "firmware_version": "auto", "ip_address": "auto" }
```

`"auto"` / blank means *work it out or say Unknown*. `role` is *Router* when the share's host is the
default gateway, otherwise *Extender*. Name, model and firmware come from here first; whatever is
still unknown is requested from the device over UPnP when it answers. Nothing is guessed.

`client_attached_to` — `"Router"`, `"Extender"` or `""`. Set it only when the PC is plugged into a
different device than the one holding the drive; the report's topology then says so.

## The test

| Key | Meaning |
|---|---|
| `interfaces` | Labels for the connections tested, e.g. `["2.5Gbps"]` or `["2.4GHz", "5GHz", "6GHz", "MLO"]`. Wired labels are descriptive only. |
| `active_connection_type` | `auto` (use what the PC is actually on), `wired`, `wireless` or `all`. |
| `skip_inactive_interfaces` | With `all`, skip interfaces that don't match the current connection. |
| `usb_formats` | File systems considered by `test_all_configured_usb_formats`. |
| `current_usb_format` | `auto` (measure it), or `FAT32` / `NTFS` / `exFAT`. |
| `iterations` | Rounds per case. 3 is the recommended minimum. |
| `test_timeout_seconds` | Stall limit for one copy: it fails only if no data moves for this long. |
| `max_plausible_mbps` | Sanity ceiling (default 1250). A round at zero or above this is an *error*, not a pass — it is a cached or local copy. |
| `accumulate_results` | Keep every run and merge them into one report. |

## Pass lines

```json
"thresholds_mbps": { "read": { "wired": 12, "wireless": 12 }, "write": { "wired": 10, "wireless": 10 } },
"thresholds_by_file_mbps": {
  "read":  { "wired": { "1": 10, "10": 25, "50": 35, "100": 40, "500": 45 } },
  "write": { "wired": { "1": 8,  "10": 20, "50": 28, "100": 30,  "500": 32 } }
},
"thresholds_note": "Example values only. Measure your own router and drive."
```

* `thresholds_mbps` is a flat minimum per direction and connection type — the fallback.
* `thresholds_by_file_mbps` (optional) sets a minimum **per file size in MB** for a connection type. A file between two listed sizes uses the next size down; one smaller than every listed size uses the smallest. Connection types not listed use the flat value.
* `thresholds_note` is printed in the reports so a reader knows where the lines came from.

The shipped values are **examples**. They depend on the router, the drive and the PC's link, so calibrate them
from your own measurements — ideally with a fast reference drive, since a slow USB stick fails write lines
because the *drive* is the limit.

## Wireless

| Key | Meaning |
|---|---|
| `auto_switch_wireless_bands` | Connect to each band in `wireless_band_ssids` in turn. |
| `wireless_band_ssids` | `{"5GHz": "MyNetwork", ...}`. Passwords are supplied at run time, never saved. |
| `band_settle_seconds` | Wait after joining a band before measuring (default 5). |

## Prompts

`prompt_for_missing_info` and `prompt_for_run_context` let the command-line run ask for missing values. They have no
effect in the GUI, which never blocks on a console prompt.
