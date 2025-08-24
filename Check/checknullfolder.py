import os

# ---- Path หลัก ----
base_path = r"I:\Shared drives\Data Management\1_Daily Operation\3. 2D & 3D files\10_WTN"
output_file = r"C:\Users\phunk\OneDrive\Desktop\empty_subfolders.txt"

# โฟลเดอร์ที่ต้องเช็ค
target_folders = {"2D", "3D", "Datasheet"}
empty_folders = []

for folder in os.listdir(base_path):
    folder_path = os.path.join(base_path, folder)
    if not os.path.isdir(folder_path):
        continue

    # หา subfolder ที่เป็น 2D,3D,Datasheet
    subfolders = [f for f in os.listdir(folder_path) if f in target_folders and os.path.isdir(os.path.join(folder_path, f))]

    for sub in subfolders:
        sub_path = os.path.join(folder_path, sub)

        # ถ้าไม่มีไฟล์และไม่มีโฟลเดอร์ย่อย => ว่างเปล่า
        if not os.listdir(sub_path):
            empty_folders.append((folder, sub))
            print(f"📂 {folder}\\{sub} ว่างเปล่า")

# บันทึกผลลง txt
with open(output_file, "w", encoding="utf-8") as f:
    if not empty_folders:
        f.write("✅ ไม่พบ 2D, 3D, Datasheet ที่ว่างเปล่า\n")
    else:
        f.write("⚠️ พบโฟลเดอร์ว่างเปล่า:\n\n")
        for folder, sub in empty_folders:
            f.write(f"{folder}\\{sub}\n")

print(f"\n✅ ตรวจสอบเสร็จ รายงานถูกบันทึกไว้ที่: {output_file}")
