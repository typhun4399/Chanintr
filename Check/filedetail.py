import os

# ---- Path หลัก ----
base_path = r"G:\Shared drives\Data Management\1_Daily Operation\3. 2D & 3D files\11_MUU - Usa checking"
output_file = r"C:\Users\tanapat\Desktop\folders_2D3D_detail.txt"

results = []

# loop ตามโฟลเดอร์หลัก
for folder in os.listdir(base_path):
    folder_path = os.path.join(base_path, folder)
    if os.path.isdir(folder_path):
        folder_data = {}

        # ---- ตรวจโฟลเดอร์ 2D ----
        path_2d = os.path.join(folder_path, "2D")
        if os.path.isdir(path_2d):
            files_2d = [f for f in os.listdir(path_2d) if os.path.isfile(os.path.join(path_2d, f))]
            folder_data["2D"] = files_2d if files_2d else None

        # ---- ตรวจโฟลเดอร์ 3D ----
        path_3d = os.path.join(folder_path, "3D")
        if os.path.isdir(path_3d):
            files_3d = [f for f in os.listdir(path_3d) if os.path.isfile(os.path.join(path_3d, f))]
            folder_data["3D"] = files_3d if files_3d else None

        # ---- ตรวจโฟลเดอร์ Datasheet ----
        path_datasheet = os.path.join(folder_path, "Datasheet")
        if os.path.isdir(path_datasheet):
            files_ds = [f for f in os.listdir(path_datasheet) if os.path.isfile(os.path.join(path_datasheet, f))]
            folder_data["Datasheet"] = files_ds if files_ds else None

        # ---- ถ้าไม่มีโฟลเดอร์ย่อยใด ๆ เลย ให้ข้าม ----
        if folder_data:
            results.append((folder, folder_data))

# ---- เขียนรายงาน ----
if results:
    with open(output_file, "w", encoding="utf-8") as f:
        f.write("📊 รายงานไฟล์ในโฟลเดอร์ 2D / 3D / Datasheet\n\n")
        for folder, data in results:
            f.write(f"{folder}\n")
            for key in ["2D", "3D", "Datasheet"]:
                if key in data:
                    files = data[key]
                    if files:
                        f.write(f"   {key}: {len(files)} ไฟล์\n")
                        for file in files:
                            f.write(f"      - {file}\n")
                    else:
                        f.write(f"   {key}: 0 ไฟล์\n")
                        f.write("      null\n")
            f.write("\n")
    print(f"✅ เสร็จสิ้น รายงานถูกบันทึกไว้ที่: {output_file}")
else:
    print("⚠️ ไม่มีโฟลเดอร์ย่อยใด ๆ ให้รายงาน")
