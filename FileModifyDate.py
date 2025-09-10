import os
from datetime import datetime
from collections import defaultdict

base_path = r"D:\AUDO\2D&3D"
log_file = r"C:\Users\tanapat\Desktop\modified_log.txt"

def get_last_modified(path):
    """คืนวันเวลาแก้ไขล่าสุดของ path"""
    if not os.path.exists(path):
        return None
    return datetime.fromtimestamp(os.path.getmtime(path))

# เก็บผลลัพธ์แบบ group by วันที่
grouped = defaultdict(list)

for folder in os.listdir(base_path):
    folder_path = os.path.join(base_path, folder)
    if not os.path.isdir(folder_path):
        continue

    for sub in ["2D", "3D", "Datasheet"]:
        sub_path = os.path.join(folder_path, sub)
        mtime = get_last_modified(sub_path)
        if mtime:
            date_str = mtime.strftime("%Y-%m-%d")  # เอาเฉพาะวันที่
            grouped[date_str].append(f"{folder}/{sub}")

# ---------------- เขียน log ลงไฟล์ ----------------
with open(log_file, "w", encoding="utf-8") as f:
    f.write("📊 การแก้ไขล่าสุดของแต่ละ subfolder Group by วันที่:\n\n")
    for date, subs in sorted(grouped.items()):
        f.write(f"📅 {date}\n")
        for s in subs:
            f.write(f"   └─ {s}\n")
        f.write("\n")

# ---------------- แสดงผลบนหน้าจอ ----------------
print(f"✅ เขียน log ลงไฟล์เรียบร้อย: {log_file}\n")
with open(log_file, "r", encoding="utf-8") as f:
    print(f.read())
