# Known limitations

Honest edges of the tool, and what to do about each.

1. **Benchmarking is Windows only.** It uses Windows APIs and tools (the WLAN API, UNC share paths, PowerShell, `netsh`). The GUI, report generation and test suite run on macOS and Linux, but the router-facing engine does not. A port would mount the share with `mount_smbfs` / `mount.cifs` and replace the Wi-Fi and link-speed probes with `networksetup` / `nmcli` equivalents.

2. **The PC's own network link limits every result.** Throughput can't exceed what the PC's adapter negotiated. A 100 Mbps link caps everything near 11 MB/s; a 1 Gbps adapter near 110 MB/s and cannot show what a 2.5G or 10G router port can do.
   *The tool reads the negotiated speed, warns below 1 Gbps (console, popup, report banner) and notes when the link is slower than the port you selected.* Use a good cable and a faster adapter to measure faster ports.

3. **The wired port is a label you choose.** Windows can't see which physical port a cable is in, so the interface name in the report is what you selected.

4. **Where the PC is connected must be declared.** A wired PC looks the same behind either the router or an extender. If your PC is on a different device than the one holding the drive, say so in *This PC connects through*; otherwise the report assumes they're the same.

5. **Pass lines are yours to calibrate.** The shipped per-file-size thresholds are examples. They depend on the router, the drive and the PC's link, and assume a fast reference drive — a slow USB stick fails the write lines because the drive is the limit.

6. **File-system detection is measured, not read.** The tool probes how the share stores timestamps to tell FAT32, exFAT and NTFS apart (the name Windows reports is often identical for every partition). It was checked against real volumes of all three. If a server ignores timestamps the result is *inconclusive* and you pick the format. HFS+ and EXT4 are not recognised by the probe; you can still set them by hand.

7. **Device details come from you or from UPnP.** Name, model and firmware are taken from the configuration first, then from the device's UPnP description if it offers one. Many routers don't, in which case the report says *Unknown*. The tool never reads a vendor's admin pages or guesses from a name.

8. **Share discovery is deliberately narrow.** It checks the default gateway and addresses that worked before — it does not sweep the network. A drive in an extender needs its address entered once; the app remembers it.

9. **Wireless band switching depends on the radios being visible.** A band is only joined if its network name is seen on that band. If all bands share one network name, the router may steer the PC to a different band than requested; the tool detects this and skips the band with the reason rather than reporting the wrong band's speed. Temporary Wi-Fi profiles created from typed passwords are removed afterwards.

10. **One drive per accumulated report.** Accumulating merges every run since the last *Clear Runs*. The report records the file system but not which physical drive was used, so keep a separate set per drive.

11. **Repeat reads can be flattered by caching.** The tool stages a fresh copy on the share for every read so the SMB client can't serve a repeat from its cache, and rejects speeds above the plausible ceiling — but a router that caches aggressively can still make reads look better than the disk.

12. **Copy timeouts are inactivity limits.** A slow but progressing copy is never cut off (hard ceiling: one hour per copy).

13. **PDF export needs Edge or Chrome.** Microsoft Edge ships with Windows 10/11; the PDF is rendered by it.

14. **2.4 GHz and some file systems were not exercised in the development lab.** The 2.4 GHz band was exercised only for its failure handling, and HFS+ / EXT4 have not been run.
