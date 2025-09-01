import os
from collections import defaultdict

# ---------------- CONFIG ----------------
base_path = r"D:\WWS\2D&3D"  # path หลัก
log_file = r"C:\Users\tanapat\Desktop\check_2D_duplicates_log.txt"

# dict เก็บ mapping -> filename : [list of folders ที่มีไฟล์นี้]
file_map = defaultdict(list)

# วนทุกโฟลเดอร์ย่อย
for root, dirs, files in os.walk(base_path):
    # เลือกเฉพาะโฟลเดอร์ที่ชื่อมี "2D" (เช่น sub2D, 2D)
    if os.path.basename(root).lower().endswith("2d"):
        for f in files:
            file_map[f].append(root)

# ---------------- SHOW & LOG RESULT ----------------
lines = []
lines.append("✅ สรุปไฟล์ที่ซ้ำกันระหว่างหลายโฟลเดอร์ 2D:\n")

duplicates_found = False
for filename, folders in file_map.items():
    if len(folders) > 1:  # มีไฟล์ซ้ำกันในหลายโฟลเดอร์
        duplicates_found = True
        lines.append(f"📂 ไฟล์: {filename}")
        for folder in folders:
            lines.append(f"   └── {folder}")
        lines.append("-" * 60)

if not duplicates_found:
    lines.append("ไม่พบไฟล์ซ้ำกันใน subfolder 2D ใด ๆ")

# เขียน log ลงไฟล์
with open(log_file, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))

print(f"\n📝 Log ถูกบันทึกที่: {log_file}")
