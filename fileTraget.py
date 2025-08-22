import os

# ---- Path หลัก ----
base_path = r"G:\Shared drives\Data Management\1_Daily Operation\3. 2D & 3D files\12_FDC"
output_file = r"C:\Users\tanapat\Desktop\folders_with_file_target.txt"

# เก็บผลลัพธ์
results = []

# เดินทุกโฟลเดอร์ย่อย
for root, dirs, files in os.walk(base_path):
    for file in files:
        if file.lower().endswith((".crdownload", ".3dm")):
            parent_main = os.path.basename(os.path.dirname(root))   # เช่น 14708_Hashira Floor Lamp
            subfolder = os.path.basename(root)                      # เช่น 2D หรือ 3D
            results.append((parent_main, subfolder, file))

# เขียนรายงาน
with open(output_file, "w", encoding="utf-8") as f:
    if not results:
        f.write("✅ ไม่พบไฟล์\n")
    else:
        f.write("🔹 รายงานโฟลเดอร์ที่มีไฟล์\n\n")
        current_folder = None
        for parent, sub, filename in sorted(results):
            if parent != current_folder:
                f.write(f"\n{parent}: ")
                current_folder = parent
            f.write(f"[{sub}]  {filename},")

print(f"✅ ตรวจสอบเสร็จ รายงานถูกบันทึกไว้ที่: {output_file}")
