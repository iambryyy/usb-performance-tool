import hashlib
import json
import os
import sys

from app_paths import app_dir

VERSION = "v1.0"

FILES = {
    "small_1mb.bin": 1,
    "document_10mb.bin": 10,
    "image_50mb.bin": 50,
    "video_100mb.bin": 100,
    "large_500mb.bin": 500,
}

MB = 1024 * 1024
SCRIPT_DIR = app_dir()
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "test_files")
MANIFEST_PATH = os.path.join(SCRIPT_DIR, "file_set_manifest.json")


def file_blocks(filename, size_mb):
    """Deterministic incompressible content: identical bytes on every machine for a given VERSION."""
    for index in range(size_mb):
        yield hashlib.shake_256(f"{VERSION}:{filename}:{index}".encode()).digest(MB)


def expected_sha256(filename, size_mb):
    digest = hashlib.sha256()
    for block in file_blocks(filename, size_mb):
        digest.update(block)
    return digest.hexdigest()


def file_sha256(path):
    digest = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(4 * MB), b""):
            digest.update(block)
    return digest.hexdigest()


def load_manifest():
    try:
        with open(MANIFEST_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return None


def write_manifest():
    files = []
    for filename, size_mb in FILES.items():
        print(f"Hashing {filename} ({size_mb} MB)...")
        files.append({"name": filename, "size_mb": size_mb, "sha256": expected_sha256(filename, size_mb)})
    manifest = {
        "version": VERSION,
        "algorithm": "sha256",
        "content": "shake_256('<version>:<name>:<block index>') x size_mb 1 MiB blocks",
        "files": files,
    }
    with open(MANIFEST_PATH, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2)
    print(f"Manifest written: {MANIFEST_PATH}")


def create_file(filename, size_mb):
    print(f"Creating {filename} ({size_mb} MB)...")
    with open(filename, "wb") as f:
        for block in file_blocks(os.path.basename(filename), size_mb):
            f.write(block)
    print(f"Done: {filename}")


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    manifest = load_manifest()
    expected = {e["name"]: e["sha256"] for e in (manifest or {}).get("files", [])}
    if manifest and manifest.get("version") != VERSION:
        print(f"Warning: manifest version {manifest.get('version')} != generator version {VERSION}; regenerate the manifest.")
        expected = {}

    for filename, size_mb in FILES.items():
        filepath = os.path.join(OUTPUT_DIR, filename)
        want = expected.get(filename) or expected_sha256(filename, size_mb)
        if os.path.exists(filepath) and os.path.getsize(filepath) == size_mb * MB and file_sha256(filepath) == want:
            print(f"OK (already matches standard set {VERSION}): {filename}")
            continue
        create_file(filepath, size_mb)

    print(f"\nAll test files match standard set {VERSION}.")


if __name__ == "__main__":
    if "--write-manifest" in sys.argv:
        write_manifest()
    else:
        main()
