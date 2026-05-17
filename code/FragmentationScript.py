import os
import shutil
import hashlib
import csv
import time
import random

# =========================
# Paths
# =========================
dataset_path = r"C:\Users\reham\Desktop\dataset2"   # Folder contains JPG, PNG, PDF, DOCX, MP4
usb_path = r"E:\\"                     # USB formatted as FAT32
csv_path = "fragmentation_log.csv"

# =========================
# Settings
# =========================
filler_folder = os.path.join(usb_path, "FILLER_FILES")
test_folder = os.path.join(usb_path, "TEST_FILES")

filler_file_size = 64 * 1024      # 64 KB filler files
number_of_fillers = 3000          # increase if USB is large
delete_ratio = 0.50               # delete 50% of filler files
delay = 0.01

# =========================
# Helper functions
# =========================
def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()

def safe_mkdir(path):
    if not os.path.exists(path):
        os.makedirs(path)

def clean_folder(path):
    if os.path.exists(path):
        shutil.rmtree(path)
    os.makedirs(path)

def create_filler_files():
    print("[1] Creating filler files...")
    clean_folder(filler_folder)

    for i in range(number_of_fillers):
        filler_path = os.path.join(filler_folder, f"filler_{i:05d}.bin")
        with open(filler_path, "wb") as f:
            f.write(os.urandom(filler_file_size))
            f.flush()
            os.fsync(f.fileno())

        if i % 100 == 0:
            print(f"Created {i} filler files")

    print("Filler files created.")

def delete_random_fillers():
    print("[2] Deleting random filler files to create free gaps...")

    files = [
        os.path.join(filler_folder, f)
        for f in os.listdir(filler_folder)
        if f.endswith(".bin")
    ]

    random.shuffle(files)
    delete_count = int(len(files) * delete_ratio)

    for i, path in enumerate(files[:delete_count]):
        os.remove(path)
        if i % 100 == 0:
            time.sleep(delay)

    print(f"Deleted {delete_count} filler files.")

def copy_dataset_files():
    print("[3] Copying original dataset files into fragmented free spaces...")
    clean_folder(test_folder)

    rows = []

    for file_name in os.listdir(dataset_path):
        src = os.path.join(dataset_path, file_name)

        if not os.path.isfile(src):
            continue

        dst = os.path.join(test_folder, file_name)

        shutil.copy2(src, dst)

        try:
            with open(dst, "ab") as f:
                f.flush()
                os.fsync(f.fileno())
        except OSError:
            pass

        size = os.path.getsize(src)
        sha = sha256_file(src)

        rows.append([file_name, size, dst, sha])

        print(f"Copied: {file_name}")

        time.sleep(delay)

    return rows

def delete_test_files():
    print("[4] Deleting test files after fragmentation...")

    for file_name in os.listdir(test_folder):
        path = os.path.join(test_folder, file_name)
        if os.path.isfile(path):
            os.remove(path)
            time.sleep(delay)

    print("Test files deleted. Now create forensic image using FTK Imager immediately.")

def write_log(rows):
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["file_name", "file_size_bytes", "usb_path_before_deletion", "sha256"])
        writer.writerows(rows)

    print(f"Log saved: {csv_path}")

# =========================
# Main
# =========================
print("WARNING: Make sure USB is formatted as FAT32 before running this script.")
print("Do not use SSD. Use USB flash drive.")
print("Run this script as Administrator if possible.")
print()

create_filler_files()
delete_random_fillers()
rows = copy_dataset_files()
write_log(rows)

input("Press Enter to delete the test files after confirming they are copied to USB...")

delete_test_files()

print()
print("DONE.")
print("Now open FTK Imager and create a forensic image of the USB immediately.")