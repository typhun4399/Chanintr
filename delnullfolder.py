import os
import shutil

# ---- Path หลัก ----
base_path = r"G:\Shared drives\Data Management\1_Daily Operation\3. 2D & 3D files\12_FDC"

# โฟลเดอร์ที่ต้องเช็ค
target_folders = {"2D", "3D", "Datasheet"}

deleted_folders = []

for folder in os.listdir(base_path):
    folder_path = os.path.join(base_path, folder)
    if not os.path.isdir(folder_path):
        continue

    # หา subfolder ที่เป็น 2D,3D,Datasheet
    subfolders = [f for f in os.listdir(folder_path) if f in target_folders and os.path.isdir(os.path.join(folder_path, f))]

    for sub in subfolders:
        sub_path = os.path.join(folder_path, sub)

        # ถ้าโฟลเดอร์นั้น "ไม่มีไฟล์และไม่มีโฟลเดอร์ย่อย" -> ลบ
        if not os.listdir(sub_path):
            shutil.rmtree(sub_path)
            deleted_folders.append(f"{folder}\\{sub}")
            print(f"🗑️ ลบ {folder}\\{sub} แล้ว (ว่างเปล่า)")

# สรุปผล
if not deleted_folders:
    print("\n✅ ไม่พบ 2D, 3D, Datasheet ที่ว่างเปล่า")
else:
    print("\n⚠️ โฟลเดอร์ที่ถูกลบ:")
    for d in deleted_folders:
        print(f" - {d}")
