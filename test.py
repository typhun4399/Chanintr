import pandas as pd
import pdfplumber
import os

# 1. กำหนด path ของไฟล์ (ใช้ path ตามที่คุณแจ้งมา)
excel_path = r"C:\Users\tanapat\Downloads\CHA STOCK (EOL) 251207 - ex Indonesia_updated 251219.xlsx"
pdf_path = r"C:\Users\tanapat\Downloads\PI confirmation for STOCK (EOL) order.pdf"
output_path = r"C:\Users\tanapat\Downloads\CHA_STOCK_Checked_Result.xlsx"


def check_stock_in_pdf():
    print("กำลังเริ่มกระบวนการ...")

    # 2. อ่านไฟล์ PDF และดึงข้อความทั้งหมดออกมา
    print(f"กำลังอ่านไฟล์ PDF: {os.path.basename(pdf_path)}")
    full_pdf_text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                full_pdf_text += page.extract_text() + "\n"
    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการอ่าน PDF: {e}")
        return

    # 3. อ่านไฟล์ Excel
    print(f"กำลังอ่านไฟล์ Excel: {os.path.basename(excel_path)}")
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการอ่าน Excel: {e}")
        return

    # ตรวจสอบว่ามีคอลัมน์ 'Commercial Code' หรือไม่
    if "Commercial Code" not in df.columns:
        print("ไม่พบคอลัมน์ 'Commercial Code' ในไฟล์ Excel โปรดตรวจสอบชื่อหัวตาราง")
        return

    # 4. ทำการค้นหา (Loop)
    print("กำลังตรวจสอบข้อมูล...")

    # ฟังก์ชันสำหรับเช็ค (ใช้ string search ธรรมดา)
    def search_in_pdf(code):
        if pd.isna(code):  # กรณีช่องว่าง
            return "No Data"

        search_term = str(code).strip()  # แปลงเป็นตัวหนังสือและตัดช่องว่างหน้าหลัง
        if search_term in full_pdf_text:
            return "Found"  # เจอ
        else:
            return "Not Found"  # ไม่เจอ

    # สร้างคอลัมน์ใหม่ชื่อ 'Status in PDF'
    df["Status in PDF"] = df["Commercial Code"].apply(search_in_pdf)

    # 5. บันทึกไฟล์ใหม่
    try:
        df.to_excel(output_path, index=False)
        print("-" * 30)
        print(f"เสร็จเรียบร้อย! บันทึกไฟล์ผลลัพธ์ที่:\n{output_path}")
        print("-" * 30)
    except Exception as e:
        print(f"ไม่สามารถบันทึกไฟล์ได้ (อาจจะเปิดไฟล์ค้างอยู่): {e}")


if __name__ == "__main__":
    check_stock_in_pdf()
