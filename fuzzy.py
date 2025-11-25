import fitz  # PyMuPDF
import pandas as pd
from rapidfuzz import fuzz

# ---------------- FILE PATHS ----------------
pdf_path = r"C:\Users\tanapat\Downloads\1_kettal_price_list_euro_25_(valid 15Sep25 - 31Aug26).pdf"
excel_path = r"C:\Users\tanapat\Desktop\vendor_code.xlsx"
output_excel = r"C:\Users\tanapat\Desktop\vendor_search_result_fuzzy.xlsx"

# ---------------- LOAD EXCEL ----------------
df = pd.read_excel(excel_path)
vendor_cols = [col for col in df.columns if col.startswith("Vendor")]

if not vendor_cols:
    raise ValueError("ไม่พบคอลัมน์ Vendor ใน Excel")

# ---------------- OPEN PDF ----------------
doc = fitz.open(pdf_path)

results = []

print("\n====== เริ่มค้นหาแบบ Fuzzy ======\n")

for idx, row in df.iterrows():
    vendors_in_row = [str(row[col]).strip() for col in vendor_cols if pd.notna(row[col])]
    if not vendors_in_row:
        continue

    print(f"🔹 Row {idx+1} ตรวจ {', '.join(vendors_in_row)}")

    for vendor in vendors_in_row:
        best_score = 0
        best_page = "-"
        for page_num, page in enumerate(doc, start=1):
            text = page.get_text() or ""
            # เปรียบเทียบ fuzzy (partial_ratio)
            score = fuzz.partial_ratio(vendor.lower(), text.lower())
            if score > best_score:
                best_score = score
                best_page = page_num

        # สมมติว่าถ้า score >= 50 ถือว่าเจอ
        found = best_score >= 50
        print(f"    Vendor: {vendor} | Best Score: {best_score} | Page: {best_page} | Found: {'Yes' if found else 'No'}")

        results.append({
            "Row": idx + 1,
            "Vendor": vendor,
            "Found": "Yes" if found else "No",
            "Page": best_page,
            "Score": best_score
        })

doc.close()

# ---------------- EXPORT EXCEL ----------------
result_df = pd.DataFrame(results)
result_df.to_excel(output_excel, index=False)
print(f"\n✅ บันทึกผลลง Excel: {output_excel}")
