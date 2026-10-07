"""
Generate a sample report from SYNTHETIC data -- no hardware needed.

Every number here is invented by a small throughput model (large files approach a
link-limited ceiling, small files are dominated by per-file overhead, writes are slower
than reads) plus a little random noise. It exists so you can see what the HTML report and
the Excel workbook look like before pointing the tool at a real router.

    python examples/make_demo_report.py
"""
import json
import os
import random
import shutil
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(os.path.dirname(HERE), "src")
OUT = os.path.join(HERE, "sample_report")

# Keep the generated files out of the project folder while the engine runs.
work = tempfile.mkdtemp(prefix="usb_tool_demo_")
sys.path.insert(0, ROOT)
import app_paths
app_paths.app_dir = lambda: work
for name in ("chart.umd.min.js", "file_set_manifest.json"):
    shutil.copy(os.path.join(ROOT, name), work)
import usb_test as u

u.LOG_PATH = "usb_test_demo.log"                       # shown in the report; no real path

random.seed(7)
CEILING = {"read": 105.0, "write": 62.0}               # MB/s a large file approaches
OVERHEAD_MB = {"read": 3.0, "write": 5.0}              # per-file cost, expressed as megabytes of lost time
FORMAT_FACTOR = {"FAT32": 0.97, "NTFS": 0.99, "exFAT": 1.00}
ROLE = {                                               # name, relative speed, share, address
    "Router": ("Living Room Router", 1.00, r"\\192.168.1.1\disk1", "192.168.1.1"),
    "Extender": ("Hallway Extender", 0.93, r"\\192.168.1.2\disk1", "192.168.1.2"),
}
FILES = (("small_1mb.bin", 1), ("document_10mb.bin", 10), ("image_50mb.bin", 50),
         ("video_100mb.bin", 100), ("large_500mb.bin", 500))
INTERFACE = "2.5Gbps"

config = json.load(open(os.path.join(ROOT, "test_config.example.json"), encoding="utf-8"))
config.update({
    "project_name": "USB Performance Test (demo data)", "output_path": OUT, "accumulate_results": False,
    "actual_connection_type": "wired", "detected_gateway_ip": "192.168.1.1", "dut_reachable": True,
    "share_write_access": True, "client_link": {"adapter": "Ethernet", "mbps": 1000.0}, "log_file": "usb_test_demo.log",
    "client_link_note": ("This PC's link (Ethernet) is 1 Gbps while the selected router port is 2.5 Gbps, so large-file "
                         "speeds can be limited to about 118 MB/s by this PC. A faster adapter is needed to measure "
                         "the router fully."),
    "file_set_summary": "Standard set v1.0: 5 of 5 file(s) used in this run verified",
    "dut": {"role": "Router", "name": "Living Room Router", "model": "ExampleCo X1", "firmware_version": "1.0.0",
            "info_source": "config"},
    "topology": "Router USB performance test - wired client",
})

results, clock = [], 0
for role, (dut_name, relative, share, address) in ROLE.items():
    for usb_format in ("FAT32", "NTFS", "exFAT"):
        for file_name, size_mb in FILES:
            for direction in ("read", "write"):
                for iteration in (1, 2, 3):
                    speed = (CEILING[direction] * size_mb / (size_mb + OVERHEAD_MB[direction])
                             * relative * FORMAT_FACTOR[usb_format] * random.gauss(1.0, 0.03))
                    if role == "Extender" and usb_format == "FAT32" and file_name == "large_500mb.bin" \
                            and direction == "write" and iteration == 2:
                        speed *= 0.45                    # one invented dip, so the failure analysis has something to show
                    speed = round(max(speed, 0.5), 2)
                    limit = u.threshold(config, direction, "wired", size_mb)
                    passed = speed >= limit
                    clock += 3
                    results.append(u.Result(
                        project_name=config["project_name"], dut_model="ExampleCo X1", dut_role=role, firmware_version="1.0.0",
                        topology=f"{role} USB performance test - wired client", client_type="Windows PC", client_os="Windows 11",
                        client_ip="192.168.1.50", configured_interface=INTERFACE, connection_type="wired",
                        actual_ssid="", actual_channel="", actual_band="", detected_gateway_ip="192.168.1.1",
                        dut_name=dut_name, detected_hardware_version="1.0", usb_share_path=share, actual_usb_format=usb_format,
                        expected_interface=INTERFACE, interface_match=True, dut_reachable=True, share_write_access=True,
                        usb_format=usb_format, direction=direction, file_set_version="v1.0", file_name=file_name,
                        iteration=iteration, throughput_mbps=speed, duration_seconds=round(size_mb / speed, 3), passed=passed,
                        status="PASS" if passed else "PERF_FAIL",
                        failure_reason="" if passed else f"throughput below configured threshold ({speed:.2f} < {limit:.2f} MB/s)",
                        timestamp=f"2026-01-15T10:{clock // 60 % 60:02d}:{clock % 60:02d}", file_verification="verified"))

if os.path.isdir(OUT):
    shutil.rmtree(OUT)
u.save_outputs(config, results)
print(f"\nDemo report written to {OUT}")
print("  rounds:", len(results), "| failed:", sum(1 for r in results if r.status != "PASS"))
shutil.rmtree(work, ignore_errors=True)
