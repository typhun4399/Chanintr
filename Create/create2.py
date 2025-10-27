import os
import pandas as pd
import re

def sanitize_folder_name(name):
    # แปลงอักขระต้องห้ามในชื่อโฟลเดอร์
    name = str(name).strip()
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    name = name.replace(' ', ' ')  # non-breaking space
    name = name.strip()
    return name

def rename_folders(base_path, excel_path):
    try:
        df = pd.read_excel(excel_path)

        if 'id' not in df.columns or 'product_title' not in df.columns:
            print("❌ ไม่พบคอลัมน์ 'id' หรือ 'product_title' ในไฟล์ Excel")
            return

        # สร้าง dictionary id -> product_title
        id_title_map = {
            str(row['id']).strip(): sanitize_folder_name(row['product_title'])
            for _, row in df.iterrows()
            if pd.notna(row['id']) and pd.notna(row['product_title'])
        }

        # เดินเข้าไปในทุกโฟลเดอร์ภายใต้ base_path
        for folder_name in os.listdir(base_path):
            old_folder_path = os.path.join(base_path, folder_name)
            if not os.path.isdir(old_folder_path):
                continue

            clean_folder_name = sanitize_folder_name(folder_name)
            folder_id = clean_folder_name.split('_')[0]  # ดึงแค่ ID ก่อน _
            
            if folder_id in id_title_map:
                new_folder_name = f"{folder_id}_{id_title_map[folder_id]}"
                new_folder_path = os.path.join(base_path, new_folder_name)

                if new_folder_path != old_folder_path:
                    if not os.path.exists(new_folder_path):
                        os.rename(old_folder_path, new_folder_path)
                        print(f"✅ เปลี่ยนชื่อ: {folder_name} → {new_folder_name}")
                    else:
                        print(f"⚠️ ชื่อเป้าหมายมีอยู่แล้ว: {new_folder_name}, ข้าม")
                else:
                    print(f"🔁 ชื่อเหมือนเดิม: {folder_name}")
            else:
                print(f"❓ ไม่พบ id '{folder_id}' ใน Excel, ข้าม")

    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาด: {e}")

# === เรียกใช้งาน ===
if __name__ == "__main__":
    base_folder_path = r"D:\ARK\2D&3D"
    excel_file_path = r"C:\Users\tanapat\Downloads\1_ARK model id to find 2D3D files_17Oct25.xlsx"

    rename_folders(base_folder_path, excel_file_path)
