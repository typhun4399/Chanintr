import os

# ---- Path หลัก ----
base_path = r"G:\Shared drives\Data Management\1_Daily Operation\3. 2D & 3D files\11_MUU - Usa checking"
output_file = r"C:\Users\tanapat\Desktop\folders_2D3D_detail.txt"

results = []

# loop ตามโฟลเดอร์หลัก
for folder in os.listdir(base_path):
    folder_path = os.path.join(base_path, folder)
    if os.path.isdir(folder_path):
        files_2d, files_3d = [], []

        # ---- ตรวจโฟลเดอร์ 2D ----
        path_2d = os.path.join(folder_path, "2D")
        if os.path.isdir(path_2d):
            files_2d = [f for f in os.listdir(path_2d) if os.path.isfile(os.path.join(path_2d, f))]

        # ---- ตรวจโฟลเดอร์ 3D ----
        path_3d = os.path.join(folder_path, "3D")
        if os.path.isdir(path_3d):
            files_3d = [f for f in os.listdir(path_3d) if os.path.isfile(os.path.join(path_3d, f))]

        results.append((folder, files_2d, files_3d))

# ---- เขียนรายงาน ----
with open(output_file, "w", encoding="utf-8") as f:
    f.write("📊 รายงานไฟล์ในโฟลเดอร์ 2D / 3D\n\n")
    for folder, f2d, f3d in results:
        f.write(f"📂 {folder}\n")
        f.write(f"   2D: {len(f2d)} ไฟล์\n")
        for file in f2d:
            f.write(f"      - {file}\n")
        f.write(f"   3D: {len(f3d)} ไฟล์\n")
        for file in f3d:
            f.write(f"      - {file}\n")
        f.write("\n")

print(f"✅ เสร็จสิ้น รายงานถูกบันทึกไว้ที่: {output_file}")
