import pandas as pd
from docx import Document
import os

# -----------------------------
# 1. ตั้งค่า path ไฟล์ input/output
# -----------------------------
input_path = r"D:\Devtest\WordToExcel\UE-35-06-009-I PCBA test for 10CS25x00020000-C-x.docx"      # <<<<< ใส่ path ของไฟล์ Word
output_path = os.path.splitext(input_path)[0] + "_converted.xlsx"

# -----------------------------
# 2. ฟังก์ชันอ่านตารางทั้งหมดใน Word
# -----------------------------
def read_docx_tables(path):
    doc = Document(path)
    all_data = []

    for table in doc.tables:
        headers = []
        first_row = table.rows[0]
        for cell in first_row.cells:
            headers.append(cell.text.strip())

        for row in table.rows[1:]:
            row_data = [cell.text.strip() for cell in row.cells]
            all_data.append(dict(zip(headers, row_data)))

    return pd.DataFrame(all_data)

# -----------------------------
# 3. แปลงข้อมูล + clean column name
# -----------------------------
df = read_docx_tables(input_path)
df.columns = [c.strip().replace("\n", " ") for c in df.columns]

# -----------------------------
# 4. mapping ชื่อคอลัมน์จาก Word → Excel
# -----------------------------
mapping = {
    'ID': 'Index',
    'Section Name': 'Name',
    'Description': 'Descriptions',
    'Procedure': 'Procedure',
    'Equipment': 'Equipment',
    'LSL': 'Low Limit',
    'Target': 'Value',
    'USL': 'High Limit',
    'Unit': 'Unit',
}

df_excel = pd.DataFrame()
for key, new_col in mapping.items():
    for col in df.columns:
        if key.lower() in col.lower():
            df_excel[new_col] = df[col]
            break
    if new_col not in df_excel.columns:
        df_excel[new_col] = ""

# -----------------------------
# 5. จัดลำดับคอลัมน์ให้ออกเหมือนเทมเพลตจริง
# -----------------------------
final_cols = [
    "Index", "Descriptions", "Equipment",
    "Low Limit", 'Value', "High Limit", "Unit"
]
df_excel = df_excel[final_cols]

# -----------------------------
# 6. บันทึกเป็น Excel
# -----------------------------
df_excel.to_excel(output_path, index=False)
print(f"✅ Done! Saved to {output_path}")
