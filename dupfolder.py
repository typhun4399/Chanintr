import os
from collections import defaultdict
from difflib import SequenceMatcher

# ---- Path หลัก ----
base_path = r"G:\Shared drives\Data Management\1_Daily Operation\3. 2D & 3D files\1_MNT"
output_file = r"C:\Users\tanapat\Desktop\similar_folders_report.txt"

# ---- เก็บ prefix -> รายชื่อ folder ----
prefix_dict = defaultdict(list)

for folder in os.listdir(base_path):
    folder_path = os.path.join(base_path, folder)
    if not os.path.isdir(folder_path):
        continue

    # แยกชื่อหลัง "_" (ถ้ามีหลาย _ เอาส่วนแรกหลัง _)
    parts = folder.split("_")
    prefix = parts[1] if len(parts) > 1 else folder
    prefix_dict[prefix].append(folder)

# ---- เช็ค exact match (prefix ซ้ำกันตรง ๆ) ----
exact_duplicates = {pre: names for pre, names in prefix_dict.items() if len(names) > 1}

# ---- เช็คความคล้ายกัน (fuzzy) โดยใช้ SequenceMatcher ----
similar_pairs = []
prefixes = list(prefix_dict.keys())
for i in range(len(prefixes)):
    for j in range(i + 1, len(prefixes)):
        ratio = SequenceMatcher(None, prefixes[i], prefixes[j]).ratio()
        if ratio >= 0.8:  # ปรับเกณฑ์ความคล้ายได้
            similar_pairs.append((prefixes[i], prefixes[j], ratio))

# ---- บันทึกผลลง txt ----
with open(output_file, "w", encoding="utf-8") as f:
    f.write("🔹 Exact duplicates (prefix ซ้ำกันตรง ๆ):\n")
    if not exact_duplicates:
        f.write("  ✅ ไม่มี prefix ซ้ำกันตรง ๆ\n\n")
    else:
        for pre, names in exact_duplicates.items():
            f.write(f"▶ {pre} ({len(names)} folders)\n")
            for folder in names:
                f.write(f"    - {folder}\n")
        f.write("\n")

    f.write("🔹 Similar folders (prefix คล้ายกัน):\n")
    if not similar_pairs:
        f.write("  ✅ ไม่มี prefix คล้ายกัน\n")
    else:
        for p1, p2, ratio in similar_pairs:
            f.write(f"  '{p1}' ~ '{p2}' (ความคล้าย {ratio:.2f})\n")

print(f"✅ ตรวจสอบเสร็จ รายงานถูกบันทึกไว้ที่: {output_file}")
