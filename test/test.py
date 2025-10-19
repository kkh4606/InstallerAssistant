import os
import shutil
import string
from collections import Counter
import rarfile
import re
import winshell
import sys

if getattr(sys, "frozen", False):
    # Running as compiled exe
    current_directory = os.path.dirname(sys.executable)
else:
    # Running as script
    current_directory = os.path.realpath(os.path.dirname(__file__))


pattern = re.compile(r"^[A-Za-z0-9 _-]+\.part\d+\.rar$")
# pattern = re.compile(r"^Bandicam\.part\d+\.rar$")


rarfile.UNRAR_TOOL = r"C:\Program Files\WinRAR\Unrar.exe"


def safe_remove_duplicates(folder_path):
    """
    Remove duplicate RAR parts safely:
    - Keeps the largest file per part number
    - Never remove a part if it's the only file for that part
      even if its size is smaller than expected (corrupted)
    """
    rar_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".rar")]
    part_groups = {}
    part_num_re = re.compile(r"(.*\.part0*(\d+))\s*(\(\d+\))?\.rar$", re.IGNORECASE)

    for f in rar_files:
        m = part_num_re.match(f)
        if not m:
            continue
        part_number = int(m.group(2))
        part_groups.setdefault(part_number, []).append(f)

    deleted_files = []

    for part_number, files in part_groups.items():
        if len(files) > 1:
            sizes = {f: os.path.getsize(os.path.join(folder_path, f)) for f in files}
            largest_size = max(sizes.values())
            largest_files = [f for f, s in sizes.items() if s == largest_size]
            keep_file = min(largest_files, key=lambda f: f.count("(") + f.count(")"))

            for f in files:
                if f != keep_file:
                    os.remove(os.path.join(folder_path, f))
                    deleted_files.append(f)

    if deleted_files:
        with open(os.path.join(folder_path, "deleted_duplicates.log"), "w") as log:
            for f in deleted_files:
                log.write(f"{f}\n")
    return deleted_files


def check_filename():
    for file in os.listdir(current_directory):

        if not file.endswith(".rar"):
            continue

        file_path = os.path.join(current_directory, file)

        if not pattern.match(file):
            removed_extra_words = re.sub(f"\s*\(\d+\)", "", file)  # type: ignore
            match_result = removed_extra_words.replace(" ", "")
            os.rename(file_path, match_result)


def verify_rar_parts(folder_path):
    """Detect potentially corrupted or missing .partXX.rar files by comparing file sizes."""
    rar_parts = [
        f
        for f in os.listdir(folder_path)
        if f.lower().endswith(".rar") and pattern.match(f)
    ]

    if not rar_parts:
        return True  # no rar files found

    # Sort parts based on part number
    part_num_re = re.compile(r"\.part0*([0-9]+)\.rar$", re.IGNORECASE)
    parsed_parts = []
    for f in rar_parts:
        m = part_num_re.search(f)
        if m:
            parsed_parts.append((int(m.group(1)), f))

    if not parsed_parts:
        print("⚠️ No valid .partXX.rar files found.")
        return False

    parsed_parts.sort()
    part_numbers = [p for p, _ in parsed_parts]
    min_part, max_part = part_numbers[0], part_numbers[-1]

    # Check for missing parts
    missing = [i for i in range(min_part, max_part + 1) if i not in part_numbers]

    # Collect file sizes
    sizes = {f: os.path.getsize(os.path.join(folder_path, f)) for _, f in parsed_parts}

    # Get expected size from most common among all except last
    non_last_parts = [f for n, f in parsed_parts if n != max_part]
    if non_last_parts:
        from collections import Counter

        expected_size = Counter(
            os.path.getsize(os.path.join(folder_path, f)) for f in non_last_parts
        ).most_common(1)[0][0]
    else:
        expected_size = list(sizes.values())[0]  # single part archive

    # Detect corrupted parts (smaller than expected, excluding last part)
    corrupted = []
    for n, f in parsed_parts:
        fsize = sizes[f]
        if n == max_part:
            continue  # last part can be smaller
        if fsize < expected_size:
            corrupted.append((f, fsize))

    if missing or corrupted:
        print("❌ Problems found with RAR parts:")
        if missing:
            print("   Missing parts:", ", ".join(f"part{i}" for i in missing))
        if corrupted:
            print("   Corrupted parts (smaller than expected):")
            for f, s in corrupted:
                print(f"     - {f} ({s} bytes)")

        with open(os.path.join(folder_path, "corrupted_files.log"), "w") as log:
            log.write(f"Expected size: {expected_size}\n")
            if missing:
                log.write("Missing parts:\n")
                for i in missing:
                    log.write(f"  part{i}\n")
            if corrupted:
                log.write("Corrupted parts:\n")
                for f, s in corrupted:
                    log.write(f"  {f} ({s} bytes)\n")
        return False

    print("✅ All .rar parts look consistent (last part may be smaller).")
    return True


def extract_rar_files(output_dir):
    safe_remove_duplicates(current_directory)
    check_filename()

    # Check for corrupted parts before extracting
    if not verify_rar_parts(current_directory):
        print("Extraction aborted due to corrupted files.")
        return

    rar_files = [f for f in os.listdir(current_directory) if pattern.match(f)]

    for file in rar_files:
        rar_path = os.path.join(current_directory, file)
        rar_file = rarfile.RarFile(rar_path)
        try:
            print(f"Please wait while extracting {file}...")
            rar_file.extractall(path=output_dir, pwd="pes26smokepatch")
            print(f"{file} → Completed ✅")

            with open(os.path.join(current_directory, "extract_success.log"), "w") as f:
                f.write("extract success")

        except Exception as e:
            print(f"Error extracting {file}: {e}")
            with open(os.path.join(current_directory, "extract_error.log"), "w") as f:
                f.write(str(e))
        finally:
            rar_file.close()


# extract_rar_files()  # type: ignore


def find_large_drive(min_free_gb=100):
    available_drives = [
        f"{d}:\\" for d in string.ascii_uppercase if os.path.exists(f"{d}:\\")
    ]

    for drive in available_drives:
        try:
            total, used, free = shutil.disk_usage(drive)
            free_gb = free / (1024**3)
            print(f"{drive} -> Free: {free_gb:.2f} GB")

            if free_gb > min_free_gb:
                print(f"✅ Drive {drive} has enough space!")
                return drive  # return the first suitable drive
        except PermissionError:
            continue  # skip drives without access

    print("❌ No drive found with enough space.")
    return None


def create_shortcut(target_exe_path):
    desktop = winshell.desktop()
    shortcut_path = os.path.join(desktop, os.path.basename(target_exe_path) + ".lnk")

    with winshell.shortcut(shortcut_path) as link:  # type: ignore
        link.path = target_exe_path
        link.description = "Shortcut to PES21"
        link.working_directory = os.path.dirname(target_exe_path)
        link.icon_location = (target_exe_path, 0)

    print(f"🎮 Shortcut created on Desktop: {shortcut_path}")


def find_exe_and_create_shortcut(root_folder):
    for root, dirs, files in os.walk(root_folder):
        for file in files:
            if file.lower().endswith(".exe"):
                exe_path = os.path.join(root, file)
                print(f"Found EXE: {exe_path}")
                create_shortcut(exe_path)
                return  # stop after first exe found


targe_drive = find_large_drive(100)

if targe_drive:
    output_dir = os.path.join(targe_drive, "GAMES")
    os.makedirs(output_dir, exist_ok=True)
    extract_rar_files(output_dir)

    if os.path.join(targe_drive, "GAMES"):
        find_exe_and_create_shortcut(os.path.join(targe_drive, "GAMES"))
    if os.path.join(targe_drive, "GAMES", "Bandicam", "KONAMI"):

        try:
            docs_patch = os.path.join(os.path.expanduser("~"), "Documents")
            if os.path.exists(os.path.join(docs_patch, "KONAMI")):  # type: ignore
                shutil.rmtree(os.path.join(docs_patch, "KONAMI"))  # type: ignore

            shutil.move(
                os.path.join(targe_drive, "GAMES", "Bandicam", "KONAMI"),
                os.path.join(os.path.expanduser("~"), "Documents"),
            )

            print("success")

        except Exception as e:
            with open(os.path.join(current_directory, "error.log"), "w") as f:
                f.write(str(e))


input("Press Enter to exit...")
