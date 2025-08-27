import pandas as pd
import os
import re # นำเข้าโมดูล re สำหรับการทำความสะอาดชื่อโฟลเดอร์

def create_folders_from_excel(excel_path, target_base_path):
    try:
        # อ่านไฟล์ Excel เข้ามาใน DataFrame
        df = pd.read_excel(excel_path)

        # ตรวจสอบว่า Base Path ที่จะสร้างโฟลเดอร์มีอยู่จริงหรือไม่ ถ้าไม่มีให้สร้าง
        if not os.path.exists(target_base_path):
            print(f"Base path '{target_base_path}' ไม่มีอยู่ กำลังสร้าง...")
            os.makedirs(target_base_path)

        # กำหนดชื่อโฟลเดอร์ย่อยที่คุณต้องการสร้าง
        subfolders_to_create = ["2D", "3D","Datasheet"]

        # วนลูปผ่านแต่ละค่าในคอลัมน์ 'FolderName'
        # .dropna() เพื่อไม่เอาค่าว่าง, .unique() เพื่อไม่สร้างโฟลเดอร์ซ้ำ
        for folder_name_raw in df['id'].dropna().unique():
            # แปลงค่า folder_name ให้เป็น string และทำความสะอาดอักขระที่ไม่ถูกต้องสำหรับชื่อโฟลเดอร์
            # ใช้ regex เพื่อลบอักขระที่ไม่ถูกต้องทั้งหมดในชื่อไฟล์ Windows
            folder_name = str(folder_name_raw).strip()
            folder_name = re.sub(r'[<>:"/\\|?*]', '_', folder_name) # แทนที่อักขระที่ไม่ถูกต้องด้วย '_'
            folder_name = folder_name.replace(' ', ' ') # แทนที่ non-breaking space ด้วย regular space
            folder_name = folder_name.strip() # ลบช่องว่างอีกครั้งหลังการแทนที่

            if not folder_name: # ข้ามถ้าชื่อว่างเปล่าหลังจากทำความสะอาด
                print(f"คำเตือน: ชื่อโฟลเดอร์ที่ทำความสะอาดแล้วจาก '{folder_name_raw}' เป็นค่าว่าง กำลังข้ามการสร้างโฟลเดอร์นี้")
                continue

            full_main_folder_path = os.path.join(target_base_path, folder_name)

            # ตรวจสอบว่าโฟลเดอร์หลักยังไม่มีอยู่ก่อนที่จะสร้าง
            if not os.path.exists(full_main_folder_path):
                try:
                    os.makedirs(full_main_folder_path)
                    print(f"สร้างโฟลเดอร์หลัก: {full_main_folder_path}")
                except OSError as e:
                    print(f"ข้อผิดพลาดในการสร้างโฟลเดอร์หลัก {full_main_folder_path}: {e}")
                    continue # ข้ามไปยังโฟลเดอร์ถัดไปหากสร้างโฟลเดอร์หลักไม่ได้
            else:
                print(f"โฟลเดอร์หลักมีอยู่แล้ว: {full_main_folder_path}")

            # สร้างโฟลเดอร์ย่อยภายในโฟลเดอร์หลัก
            for subfolder in subfolders_to_create:
                full_subfolder_path = os.path.join(full_main_folder_path, subfolder)
                if not os.path.exists(full_subfolder_path):
                    try:
                        os.makedirs(full_subfolder_path)
                        print(f"  - สร้างโฟลเดอร์ย่อย: {full_subfolder_path}")
                    except OSError as e:
                        print(f"  - ข้อผิดพลาดในการสร้างโฟลเดอร์ย่อย {full_subfolder_path}: {e}")
                else:
                    print(f"  - โฟลเดอร์ย่อยมีอยู่แล้ว: {full_subfolder_path}")

    except FileNotFoundError:
        print(f"ข้อผิดพลาด: ไม่พบไฟล์ Excel ที่ '{excel_path}'")
    except Exception as e:
        print(f"เกิดข้อผิดพลาดที่ไม่คาดคิด: {e}")

# --- กำหนด Path ของไฟล์ Excel และ Path ปลายทางของคุณ ---
if __name__ == "__main__":
    # พาธของไฟล์ Excel ที่มีคอลัมน์ 'FolderName'
    excel_file_path = r"C:\Users\tanapat\Downloads\1_VCT & VCC_id no to review 2D-3D_Aug25.xlsx"
    
    # พาธของไดเรกทอรีหลักที่คุณต้องการให้โฟลเดอร์ใหม่ถูกสร้างขึ้น
    destination_base_path = r"D:\VCT\2D&3D"

    # เรียกใช้ฟังก์ชันเพื่อสร้างโฟลเดอร์
    create_folders_from_excel(excel_file_path, destination_base_path)
    print("\nการสร้างโฟลเดอร์เสร็จสมบูรณ์")