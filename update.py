import os
import pandas as pd

def update_excel_based_on_2d_3d_datasheet(
    base_folder_path, excel_file_path, output_excel_file_path=None
):
    if output_excel_file_path is None:
        output_excel_file_path = excel_file_path

    # อ่านไฟล์ Excel
    try:
        df = pd.read_excel(excel_file_path)
        print(f"✅ อ่านไฟล์ Excel '{excel_file_path}' สำเร็จ")
    except FileNotFoundError:
        print(f"❌ ไม่พบไฟล์ Excel ที่ '{excel_file_path}'")
        return
    except Exception as e:
        print(f"❌ ข้อผิดพลาดในการอ่าน Excel: {e}")
        return

    # ตรวจสอบคอลัมน์ที่จำเป็น
    status_col = "2D/ 3D/ Datasheet"
    notes_col = "Notes"
    required_cols = ["id", status_col, notes_col]
    for col in required_cols:
        if col not in df.columns:
            print(f"⚠️ ไม่พบคอลัมน์ '{col}' ใน Excel")
            return

    # กำหนดค่าเริ่มต้น
    if not df.empty:
        df[status_col] = "N"
        df[notes_col] = ""
        print(f"➡️ รีเซ็ตค่า '{status_col}' = 'N', '{notes_col}' = ''")
    else:
        print("⚠️ DataFrame ว่าง")
        return

    # เตรียม id ให้พร้อมเทียบกับชื่อโฟลเดอร์
    df["id_str"] = df["id"].astype(str).str.strip()

    # ดึงรายชื่อโฟลเดอร์ใน path
    subdirectories = [
        d for d in os.listdir(base_folder_path)
        if os.path.isdir(os.path.join(base_folder_path, d))
    ]

    def has_files(path):
        return os.path.exists(path) and any(os.path.isfile(os.path.join(path, f)) for f in os.listdir(path))

    for sub_folder in subdirectories:
        folder_path = os.path.join(base_folder_path, sub_folder)
        idxs = df[df["id_str"] == sub_folder].index

        if not idxs.empty:
            print(f"➡️ ตรวจสอบโฟลเดอร์ '{sub_folder}'")

            path_2d = os.path.join(folder_path, "2D")
            path_3d = os.path.join(folder_path, "3D")
            path_ds = os.path.join(folder_path, "Datasheet")

            has_2d = has_files(path_2d)
            has_3d = has_files(path_3d)
            has_ds = has_files(path_ds)

            if has_2d and has_3d and has_ds:
                print("   ✅ ครบ 2D / 3D / Datasheet")
                for i in idxs:
                    df.at[i, status_col] = "Y"
            else:
                missing = []
                if not has_2d: missing.append("ไม่มี2D")
                if not has_3d: missing.append("ไม่มี3D")
                if not has_ds: missing.append("ไม่มีDatasheet")
                note = "/".join(missing)
                print(f"   ❌ ขาด: {note}")
                for i in idxs:
                    df.at[i, notes_col] = note
        else:
            print(f"   ⚠️ ไม่พบโฟลเดอร์สำหรับ ID '{sub_folder}' ใน Excel")

    df.drop(columns=["id_str"], inplace=True)

    try:
        df.to_excel(output_excel_file_path, index=False)
        print(f"\n✅ บันทึกเรียบร้อย: {output_excel_file_path}")
    except Exception as e:
        print(f"\n❌ ไม่สามารถบันทึกไฟล์: {e}")

# 🔧 เรียกใช้งาน
base_folder_path = r"D:\AUDO\2D&3D - Copy"
excel_file_path = r"C:\Users\tanapat\Downloads\1_AUD active model id_25Jun25.xlsx"
update_excel_based_on_2d_3d_datasheet(base_folder_path, excel_file_path)
