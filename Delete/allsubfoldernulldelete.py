import os
import shutil

base_folder_path = r"G:\Shared drives\Data Management\1_Daily Operation\3. 2D & 3D files\18_WWS_Datasheet ที่ไม่ load ขึ้น app by model id"

def has_files(path):
    """เช็คว่ามีไฟล์ในโฟลเดอร์หรือไม่"""
    return os.path.exists(path) and any(
        os.path.isfile(os.path.join(path, f)) for f in os.listdir(path)
    )

# วนเช็คแต่ละโฟลเดอร์
for folder_name in os.listdir(base_folder_path):
    folder_path = os.path.join(base_folder_path, folder_name)
    if not os.path.isdir(folder_path):
        continue

    path_2d = os.path.join(folder_path, "2D")
    path_3d = os.path.join(folder_path, "3D")
    path_ds = os.path.join(folder_path, "Datasheet")

    # ถ้าไม่มีไฟล์ 2D, 3D, Datasheet เลย → ลบโฟลเดอร์
    if not (has_files(path_2d) or has_files(path_3d) or has_files(path_ds)):
        try:
            shutil.rmtree(folder_path)
            print(f"🗑️ ลบโฟลเดอร์ว่าง '{folder_name}'")
        except Exception as e:
            print(f"❌ ไม่สามารถลบโฟลเดอร์ '{folder_name}': {e}")
