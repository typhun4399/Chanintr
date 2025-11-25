# -*- coding: utf-8 -*-
import os
import pandas as pd
import urllib.request
import time
from PIL import Image, ImageOps

# ---------------- CONFIG ----------------
excel_files = [r"C:\Users\tanapat\Downloads\base_feature_images_all_pages_Finish.xlsx"]
base_folder = r"D:\HIC Feture\Finish"
origin_folder = os.path.join(base_folder, "origin")
crop_folder = os.path.join(base_folder, "crop")

# สร้างโฟลเดอร์
os.makedirs(origin_folder, exist_ok=True)
os.makedirs(crop_folder, exist_ok=True)

initial_size = 500  # ขนาด 500x500 หลังโหลด
crop_top_bottom = 75  # ตัดบนและล่างออก 75px
crop_left_right = 75  # ตัดซ้าย/ขวาออก 75px

# ---------------- READ EXCEL ----------------
all_dfs = []
for file in excel_files:
    try:
        df = pd.read_excel(file)
        all_dfs.append(df)
        print(f"Loaded {file} with {len(df)} rows.")
    except Exception as e:
        print(f"❌ Failed to load {file}: {e}")

combined_df = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()

# ---------------- STEP 1: DOWNLOAD & RESIZE 500x500 (เก็บใน origin) ----------------
if 'image_url' not in combined_df.columns or 'name' not in combined_df.columns:
    print("❌ Columns 'image_url' or 'name' not found in Excel.")
else:
    for idx, row in combined_df.dropna(subset=['image_url', 'name']).iterrows():
        try:
            url = row['image_url']

            # ตั้งชื่อไฟล์: ถ้า code มีค่าใช้ code, ถ้าไม่มีใช้ name
            if 'code' in row and pd.notna(row['code']):
                base_name = str(row['code']).strip()
            else:
                base_name = str(row['name']).strip()

            base_name = base_name.replace('/', '_').replace('\\', '_')
            filename = f"{base_name}.jpg"
            origin_path = os.path.join(origin_folder, filename)
            crop_path = os.path.join(crop_folder, filename)

            # ดาวน์โหลดภาพ
            print(f"Downloading {idx+1}: {url} -> {filename}")
            urllib.request.urlretrieve(url, origin_path)
            time.sleep(0.2)

            # Resize เป็น 500x500
            with Image.open(origin_path) as img:
                img_resized = ImageOps.fit(img, (initial_size, initial_size), method=Image.LANCZOS)
                img_resized.save(origin_path, format='JPEG')

            # ก็อปปี้ไปโฟลเดอร์ crop สำหรับขั้นตอนต่อไป
            img_resized.save(crop_path, format='JPEG')

        except Exception as e:
            print(f"❌ Failed to process {url}: {e}")

    print("✅ Finished downloading and resizing all images to 500x500 px (saved in origin & crop)")

# ---------------- STEP 2: CROP บน/ล่าง 75px ----------------
for file in os.listdir(crop_folder):
    if file.lower().endswith(".jpg"):
        filepath = os.path.join(crop_folder, file)
        try:
            with Image.open(filepath) as img:
                width, height = img.size
                top = crop_top_bottom
                bottom = height - crop_top_bottom
                img_cropped = img.crop((0, top, width, bottom))
                img_cropped.save(filepath, format='JPEG')
        except Exception as e:
            print(f"❌ Failed to crop top/bottom {file}: {e}")

print(f"✅ Finished cropping top/bottom {crop_top_bottom}px for all images in 'crop'")

# ---------------- STEP 3: CROP ซ้าย/ขวา 75px ----------------
for file in os.listdir(crop_folder):
    if file.lower().endswith(".jpg"):
        filepath = os.path.join(crop_folder, file)
        try:
            with Image.open(filepath) as img:
                width, height = img.size
                left = crop_left_right
                right = width - crop_left_right
                img_cropped = img.crop((left, 0, right, height))
                img_cropped.save(filepath, format='JPEG')
        except Exception as e:
            print(f"❌ Failed to crop left/right {file}: {e}")

print(f"✅ Finished cropping left/right {crop_left_right}px for all images in 'crop'")
