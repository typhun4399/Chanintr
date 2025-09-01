import os
from collections import defaultdict
from difflib import SequenceMatcher

# ---- Path หลัก ----
base_path = r"G:\Shared drives\Data Management\1_Daily Operation\3. 2D & 3D files\2_AUD_done uploaded"
output_file = r"C:\Users\tanapat\Desktop\similar_folders_report_with_files.txt"

# ---- ฟังก์ชันตรวจสอบความคล้าย ----
def is_similar(a, b, threshold=0.8):
    return SequenceMatcher(None, a.lower(), b.lower()).ratio() >= threshold

# ---- เก็บ prefix -> รายชื่อ folder ----
prefix_dict = defaultdict(list)

for folder in os.listdir(base_path):
    folder_path = os.path.join(base_path, folder)
    if not os.path.isdir(folder_path):
        continue

    parts = folder.split("_")
    prefix = parts[1] if len(parts) > 1 else folder
    prefix_dict[prefix].append(folder)

# ---- exact duplicates ----
exact_duplicates = {pre: names for pre, names in prefix_dict.items() if len(names) > 1}

# ---- fuzzy similarity ----
similar_pairs = []
prefixes = list(prefix_dict.keys())
for i in range(len(prefixes)):
    for j in range(i + 1, len(prefixes)):
        ratio = SequenceMatcher(None, prefixes[i], prefixes[j]).ratio()
        if ratio >= 0.8:
            similar_pairs.append((prefixes[i], prefixes[j], ratio))

# ---- ตรวจสอบไฟล์ใน 2D, 3D ----
def check_files_against_prefix(folder_path, prefix):
    result = defaultdict(list)
    for sub in ["2D", "3D"]:
        sub_path = os.path.join(folder_path, sub)
        if os.path.exists(sub_path) and os.path.isdir(sub_path):
            for file in os.listdir(sub_path):
                file_path = os.path.join(sub_path, file)
                if os.path.isfile(file_path) and is_similar(prefix, os.path.splitext(file)[0]):
                    result[sub].append(file)
    return result

# ---- เขียนรายงาน ----
with open(output_file, "w", encoding="utf-8") as f:
    f.write("🔹 Exact duplicates (prefix ซ้ำกันตรง ๆ) และไฟล์ใน 2D/3D ที่คล้ายกับ prefix:\n")
    if not exact_duplicates:
        f.write("  ✅ ไม่มี prefix ซ้ำกันตรง ๆ\n\n")
    else:
        for pre, folders in exact_duplicates.items():
            f.write(f"▶ Prefix '{pre}' ({len(folders)} folders)\n")
            for folder in folders:
                f.write(f"    - {folder}\n")
                folder_path = os.path.join(base_path, folder)
                file_check = check_files_against_prefix(folder_path, pre)
                for sub, files in file_check.items():
                    if files:
                        f.write(f"      {sub} files similar to prefix: {files}\n")

    f.write("\n🔹 Similar folders (prefix คล้ายกัน) และไฟล์ใน 2D/3D ที่คล้ายกับ prefix:\n")
    if not similar_pairs:
        f.write("  ✅ ไม่มี prefix คล้ายกัน\n")
    else:
        for p1, p2, ratio in similar_pairs:
            f.write(f"'{p1}' ~ '{p2}' (ความคล้าย {ratio:.2f})\n")
            for pre in [p1, p2]:
                for folder in prefix_dict[pre]:
                    folder_path = os.path.join(base_path, folder)
                    file_check = check_files_against_prefix(folder_path, pre)
                    for sub, files in file_check.items():
                        if files:
                            f.write(f"  - {folder} [{sub}]: {files}\n")

print(f"✅ ตรวจสอบเสร็จ รายงานถูกบันทึกไว้ที่: {output_file}")
