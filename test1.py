import os
import zipfile
import shutil
from send2trash import send2trash  # ✅ ใช้สำหรับย้ายไปถังขยะ

# 🗂 path หลัก
base_path = r"D:\MNT"

# 🔍 วนลูปทุกโฟลเดอร์ย่อยใน path
for root, dirs, files in os.walk(base_path):
    if "_MACOSX" in root:
        continue

    for file in files:
        if file.endswith(".zip"):
            zip_path = os.path.join(root, file)
            extract_folder = root
            print(f"📦 พบไฟล์ ZIP: {zip_path}")

            try:
                # 🔓 แตกไฟล์ทั้งหมด ยกเว้น _MACOSX/
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    for member in zip_ref.namelist():
                        if member.startswith("_MACOSX/"):
                            continue
                        zip_ref.extract(member, extract_folder)

                # 🗑️ ย้ายไฟล์ ZIP ไปถังขยะ
                send2trash(zip_path)
                print(f"♻️ ย้าย ZIP ไปถังขยะเรียบร้อย: {zip_path}")

                # 🧹 ถ้ามีโฟลเดอร์ _MACOSX หลังแตกไฟล์ ให้ย้ายไปถังขยะด้วย
                macosx_path = os.path.join(extract_folder, "_MACOSX")
                if os.path.exists(macosx_path):
                    send2trash(macosx_path)
                    print(f"🧹 ย้ายโฟลเดอร์ _MACOSX ไปถังขยะ: {macosx_path}")

            except Exception as e:
                print(f"❌ เกิดข้อผิดพลาดกับ {zip_path}: {e}")
