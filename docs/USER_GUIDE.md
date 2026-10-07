# User guide

## Before you start

* A Windows 10/11 PC and a router (or NAS) with a USB drive plugged in and **USB / file sharing switched on**.
* Confirm the share opens in File Explorer (for example `\\192.168.1.1\disk1`).
* For the most meaningful numbers, connect the PC with a **gigabit-or-better** link. A 100 Mbps link caps every result near 11 MB/s no matter how fast the router or drive is — the tool warns you when it sees this.

## The window

| Area | What it is for |
|---|---|
| **USB Share Path** | The share to test. The app looks for it on launch (the default gateway and hosts that worked before). Choose *All share paths* to test every partition of the drive in one run. **Scan for Shares** searches again after you move the drive. |
| **Device Role** | *Router* if the share lives on the network's gateway, *Extender* for any other device. It is filled in automatically ("detected") and you can change it. |
| **USB Format** | FAT32 / NTFS / exFAT. Filled in by measuring the drive. *Unknown* means "detect it for me". |
| **Iterations / Timeout** | Rounds per test case (3 is a sensible minimum for averages) and how long a copy may stall before it is declared failed. |
| **Accumulate Results Across Runs** | Keeps every run and merges them into one report — tick it when you want the router and the extender in the same report. |
| **Device (optional)** | Name, model and firmware to show in the report. Left blank, the tool asks the device over UPnP and otherwise says *Unknown* — it never guesses. |
| **Connection Type** | *Wireless*, or the speed of the wired port you plugged into (a label for the report — Windows can't see which port a cable is in). |
| **This PC connects through** | Leave on the first choice when your PC and the drive are on the same device. Pick the other device if, say, the drive is in the router but your cable goes to an extender. |
| **Test Files** | The files on **this PC** that the test copies to the drive and back. They are deleted from the drive again after each round. *Generate Standard Test Files* creates the fixed set (about 660 MB). |

Under the console, one line shows the size of the run you're about to start (rounds, data moved and, after your first run, a time estimate).

## Running a test

1. Check the plan line, then press **▶ Start Test**.
2. The progress bar shows rounds done and time left; the console shows what is happening. On a small screen the page scrolls — it jumps to the console when you press Start.
3. **■ Stop** ends the run at any time. Rounds already measured are kept and the report is still written; anything unfinished is marked *SKIPPED*, so a partial run is never shown as a complete pass.
4. When it finishes: **Open Report** (HTML), **Open Excel** (dashboard + data) or **Save as PDF**. These buttons stay greyed out until there is something to open.

### Wireless

Choose *Wireless*, tick the bands to test (2.4 GHz / 5 GHz / 6 GHz / MLO), and enter each band's network name and password (or tick *Same password for all*; *Show passwords* lets you check what you typed). The tool connects to each band in turn, confirms it landed on the right one, waits a few seconds for the link to settle, then measures. Passwords are used for that run only and never written to disk; network names are remembered.

If a band can't be joined it is skipped, recorded with the reason, and listed in a summary when the run ends.

### Testing a router and an extender in one report

1. Tick **Accumulate Results Across Runs**.
2. Test the drive in the router (PC on the router). 
3. Move the drive and the PC to the extender, press *Scan for Shares*, and run again with the **same port label**.
4. The report now shows both devices, both topologies and a Router vs Extender chart.

Use one drive per accumulated report — the report records the file system but not which physical drive was used. **Clear Runs** starts fresh.

## Reading the results

* **PASS / PERF_FAIL / ERROR / SKIPPED** — a round that ran and met its pass line; ran but was too slow; couldn't complete (timeout, unreachable, full drive, impossible speed); or wasn't tested (band couldn't be joined, run stopped).
* **Pass lines** are set per file size (see [CONFIGURATION.md](CONFIGURATION.md)). Small files are slower because per-file overhead dominates, so one flat number would be unfair to them.
* **Network link** (Client PC card) is what the PC's adapter negotiated. If it is below the speed of the router port you picked, the report says speeds may be limited by the PC.
* The workbook's **Results** sheet is filterable; **Issues** lists only failed, skipped and abnormal rounds; **Analysis** holds the tables behind the charts; **Notes** explains the colours and formulas.
* `USB_Performance_Results_TEMPLATE.xlsx` is an empty workbook with the same layout, for filling in by hand.

## Output files (in `test_results/`)

| File | Contents |
|---|---|
| `report.html` | Summary, environment, pass/fail verdict, charts, failures, raw data. Self-contained; works offline. |
| `usb_performance_results.xlsx` | The Excel dashboard. If it is open in Excel when a run finishes, a time-stamped copy is written beside it instead. |
| `test_results.csv` / `.json` | One row per round with every recorded field. |
| `runs/` | One file per run when accumulation is on. |
| `../logs/` | A short trail of each run (what was detected, band switches, each result). |

## Troubleshooting

* **Speeds stuck near 10–11 MB/s on wired:** the PC's link is 100 Mbps. Try another cable or port and check the adapter's *Speed & Duplex* is on *Auto Negotiation*.
* **"No USB storage found":** check the drive is plugged in and shared, then press *Scan for Shares*. If the router asks for a login, the app shows a sign-in box.
* **A band shows "Could not connect":** the message explains why. Common causes: wrong password, the band is off, or all bands share one network name and the router steered the PC to another band — give each band its own name while testing.
* **Everything fails with "Client IP not obtained":** the PC has no usable address on that network (self-assigned 169.254.x.x). Reconnect and try again.
