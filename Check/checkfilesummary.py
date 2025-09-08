import os
from collections import Counter

# ---- Path หลัก ----
base_path = r"G:\Shared drives\Data Management\1_Daily Operation\3. 2D & 3D files\18_WWS"
output_file = r"C:\Users\tanapat\Desktop\file_summary.txt"

def count_file_types(base_path):
    summary = {"2D": Counter(), "3D": Counter(), "Datasheet": Counter()}

    # loop ทุก folder ชั้นแรก
    for folder in os.listdir(base_path):
        folder_path = os.path.join(base_path, folder)
        if not os.path.isdir(folder_path):
            continue

        for sub in ["2D","3D","Datasheet"]:
            sub_path = os.path.join(folder_path, sub)
            if os.path.isdir(sub_path):
                # loop ไฟล์ใน sub folder
                for root, dirs, files in os.walk(sub_path):
                    for file in files:
                        ext = os.path.splitext(file)[1].lower()  # นามสกุลไฟล์
                        summary[sub][ext] += 1

    return summary


# ---- รันตรวจสอบ ----
results = count_file_types(base_path)

with open(output_file, "w", encoding="utf-8") as f:
    for sub in ["2D","3D","Datasheet"]:
        f.write(f"📂 สรุปไฟล์ใน {sub} รวมทุกโฟลเดอร์\n")
        if not results[sub]:
            f.write("   (ไม่มีไฟล์เลย)\n\n")
        else:
            for ext, count in results[sub].items():
                f.write(f"   {ext} : {count} ไฟล์\n")
            f.write("\n")

print(f"✅ ตรวจสอบเสร็จ ผลลัพธ์ถูกบันทึกไว้ที่: {output_file}")
