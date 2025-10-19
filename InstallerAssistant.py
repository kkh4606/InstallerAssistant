import os
import shutil
import string
import re
import winshell
import sys
import subprocess
from datetime import datetime

# --- Directory setup ---
# if getattr(sys, "frozen", False):
#     current_directory = os.path.dirname(sys.executable)
# else:
#     current_directory = os.path.realpath(os.path.dirname(__file__))
current_directory = "E:\\GAMES"
PASSWORD = "pes26smokepatch"
WINRAR_PATH = r"C:\Program Files\WinRAR\WinRAR.exe"
LOG_FILE = os.path.join(current_directory, "extract.log")
pattern = re.compile(r"^[A-Za-z0-9 _-]+\.part\d+\.rar$")


# --- Logging ---
def log(message):
    ts = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    print(f"{ts} {message}")
    with open(LOG_FILE, "a") as f:
        f.write(f"{ts} {message}\n")


# --- Disk check ---
def find_large_drive(min_free_gb=100):
    for d in string.ascii_uppercase:
        drive = f"{d}:\\"
        if not os.path.exists(drive):
            continue
        try:
            total, used, free = shutil.disk_usage(drive)
            free_gb = free / (1024**3)
            log(f"{drive} -> Free: {free_gb:.2f} GB")
            if free_gb >= min_free_gb:
                log(f"✅ Drive {drive} has enough space")
                return drive
        except PermissionError:
            continue
    log("❌ No drive found with enough space")
    return None


# --- Duplicate removal ---
def safe_remove_duplicates(folder):
    rar_files = [f for f in os.listdir(folder) if f.lower().endswith(".rar")]
    part_groups = {}
    part_num_re = re.compile(r"(.*\.part0*(\d+))\s*(\(\d+\))?\.rar$", re.IGNORECASE)
    for f in rar_files:
        m = part_num_re.match(f)
        if not m:
            continue
        part_number = int(m.group(2))
        part_groups.setdefault(part_number, []).append(f)

    deleted = []
    for part, files in part_groups.items():
        if len(files) <= 1:
            continue
        sizes = {f: os.path.getsize(os.path.join(folder, f)) for f in files}
        largest_size = max(sizes.values())
        keep_file = min(
            [f for f, s in sizes.items() if s == largest_size],
            key=lambda f: f.count("(") + f.count(")"),
        )
        for f in files:
            if f != keep_file:
                os.remove(os.path.join(folder, f))
                deleted.append(f)
    if deleted:
        with open(os.path.join(folder, "deleted_duplicates.log"), "w") as logf:
            for f in deleted:
                logf.write(f"{f}\n")
    return deleted


# --- Filename cleanup ---
def check_filename():
    for f in os.listdir(current_directory):
        if not f.lower().endswith(".rar"):
            continue
        if not pattern.match(f):
            new_name = re.sub(r"\s*\(\d+\)", "", f).replace(" ", "")
            os.rename(
                os.path.join(current_directory, f),
                os.path.join(current_directory, new_name),
            )


# --- Verify parts ---
def verify_rar_parts(folder):
    rar_parts = [
        f for f in os.listdir(folder) if f.lower().endswith(".rar") and pattern.match(f)
    ]
    if not rar_parts:
        return True
    part_num_re = re.compile(r"\.part0*([0-9]+)\.rar$", re.IGNORECASE)
    parsed = [
        (int(part_num_re.search(f).group(1)), f)  # type: ignore
        for f in rar_parts
        if part_num_re.search(f)
    ]
    if not parsed:
        log("⚠️ No valid .partXX.rar files found")
        return False
    parsed.sort()
    numbers = [p for p, _ in parsed]
    min_part, max_part = numbers[0], numbers[-1]
    missing = [i for i in range(min_part, max_part + 1) if i not in numbers]

    sizes = {f: os.path.getsize(os.path.join(folder, f)) for _, f in parsed}
    non_last = [f for n, f in parsed if n != max_part]
    expected_size = (
        max(
            set(sizes[f] for f in non_last), key=lambda x: list(sizes.values()).count(x)
        )
        if non_last
        else list(sizes.values())[0]
    )
    corrupted = [
        (f, sizes[f]) for n, f in parsed if n != max_part and sizes[f] < expected_size
    ]

    if missing or corrupted:
        log("❌ Problems found with RAR parts")
        if missing:
            log("   Missing: " + ",".join(f"part{i}" for i in missing))
        if corrupted:
            log("   Corrupted: " + ",".join(f"{f}({s})" for f, s in corrupted))
        with open(os.path.join(folder, "corrupted_files.log"), "w") as l:
            if missing:
                l.write("Missing: " + ",".join(f"part{i}" for i in missing) + "\n")
            if corrupted:
                l.write(
                    "Corrupted: " + ",".join(f"{f}({s})" for f, s in corrupted) + "\n"
                )
        return False
    log("✅ All .rar parts consistent")
    return True


# --- Extraction with progress ---
def extract_rar_files(output_dir):
    safe_remove_duplicates(current_directory)
    check_filename()
    if not verify_rar_parts(current_directory):
        log("Extraction aborted due to corrupted files")
        return

    rar_parts = sorted(
        [
            f
            for f in os.listdir(current_directory)
            if f.lower().endswith(".rar") and pattern.match(f)
        ]
    )
    if not rar_parts:
        log("❌ No RAR parts found")
        return
    if not os.path.exists(WINRAR_PATH):
        log("❌ WinRAR not found")
        return

    for part in rar_parts:
        part_path = os.path.join(current_directory, part)
        log(f"Please wait while extracting {part}...")
        # Run WinRAR in console mode
        process = subprocess.Popen(
            [WINRAR_PATH, "x", "-y", f"-p{PASSWORD}", part_path, output_dir],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            universal_newlines=True,
        )
        for line in process.stdout:  # type: ignore
            if "%" in line:
                print(line.strip())  # shows percentage in real-time
        process.wait()
        log(f"{part} → Completed ✅")


# --- Shortcut ---
def create_shortcut(target_exe):
    desktop = winshell.desktop()
    sc = os.path.join(desktop, os.path.basename(target_exe) + ".lnk")
    with winshell.shortcut(sc) as link:  # type: ignore
        link.path = target_exe
        link.description = "Shortcut to PES21"
        link.working_directory = os.path.dirname(target_exe)
        link.icon_location = (target_exe, 0)
    log(f"🎮 Shortcut created: {sc}")


def find_exe_and_create_shortcut(root):
    for r, d, f in os.walk(root):
        for file in f:
            if file.lower().endswith(".exe"):
                create_shortcut(os.path.join(r, file))
                return


# --- Main ---
drive = find_large_drive(100)
if drive:
    out_dir = os.path.join(drive, "GAMES")
    os.makedirs(out_dir, exist_ok=True)
    extract_rar_files(out_dir)
    find_exe_and_create_shortcut(out_dir)

    konami_src = os.path.join(out_dir, "Bandicam", "KONAMI")
    konami_dst = os.path.join(os.path.expanduser("~"), "Documents", "KONAMI")
    if os.path.exists(konami_src):
        if os.path.exists(konami_dst):
            shutil.rmtree(konami_dst)
        shutil.move(konami_src, konami_dst)
        log("KONAMI folder moved successfully ✅")

log("All operations completed.")
input("Press Enter to exit...")
