import os

# ---- Path หลัก ----
base_path = r"G:\Shared drives\Data Management\1_Daily Operation\3. 2D & 3D files\12_FDC"

# นามสกุลไฟล์ที่ต้องลบ
remove_exts = {".crdownload", ".3dm"}  # "" = ไม่มีนามสกุล

deleted_files = []

for folder in os.listdir(base_path):
    folder_path = os.path.join(base_path, folder)
    if not os.path.isdir(folder_path):
        continue

    sub_path = os.path.join(folder_path, "2D")
    if os.path.isdir(sub_path):
        for root, dirs, files in os.walk(sub_path):
            for file in files:
                ext = os.path.splitext(file)[1].lower()
                if ext in remove_exts:
                    file_path = os.path.join(root, file)
                    try:
                        os.remove(file_path)
                        deleted_files.append(file_path)
                        print(f"🗑 ลบแล้ว: {file_path}")
                    except Exception as e:
                        print(f"❌ ลบไม่ได้: {file_path} ({e})")

print(f"\n✅ เสร็จสิ้น ลบไฟล์ {len(deleted_files)}")
