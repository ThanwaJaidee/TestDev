import pandas as pd
import json
import os

# -----------------------------
# 1. ตั้งค่า path ไฟล์ input/output
# -----------------------------
input_path = r"D:\Devtest\ToJsonl\10002169_Increase 90 Volt.xls"
output_path = os.path.splitext(input_path)[0] + ".jsonl"  # จะได้ชื่อเดียวกันแต่สกุล .jsonl

# -----------------------------
# 2. อ่านไฟล์ Excel
# -----------------------------
df = pd.read_excel(input_path)

# ล้างชื่อคอลัมน์ (ตัดช่องว่าง ซ่อน)
df.columns = df.columns.str.strip()

# -----------------------------
# 3. ตรวจสอบชื่อคอลัมน์
# -----------------------------
def find_col(df, keywords):
    for col in df.columns:
        if any(k.lower() in str(col).lower() for k in keywords):
            return col
    return None

step_col = find_col(df, ["step"])
desc_col = find_col(df, ["desc"])
unit_col = find_col(df, ["unit"])
result_col = find_col(df, ["result", "status", "output"])

print("🧾 Detected columns:", step_col, desc_col, unit_col, result_col)

# ถ้าชื่อคอลัมน์ไม่ตรง สามารถเปลี่ยนชื่อได้เช่น
# df.rename(columns={"Step#": "Step #"}, inplace=True)

# -----------------------------
# 4. รวมข้อความเป็น instruction
# -----------------------------
df["instruction"] = (
    "Test " + df["Test #"].astype(str) + " : " +
    "Step " + df["Step #"].astype(str) + " : " +
    "Description " + df["Description"].fillna("") +
    " | Unit: " + df["Unit"].fillna("")
)

# ถ้าต้องการเพิ่มช่องอื่น เช่น “Low Limit” หรือ “High Limit” ก็ได้ เช่น
# df["instruction"] = (
#     "Step " + df["Step #"].astype(str) + ": " +
#     df["Description"].fillna("") +
#     " | Range: " + df["Low Limit"].astype(str) + " - " + df["High Limit"].astype(str)
# )

# -----------------------------
# 5. สร้างฟิลด์สำหรับ JSONL
# -----------------------------
records = []
for _, row in df.iterrows():
    record = {
        "instruction": row["instruction"],
        "output": str(row["Result"]) if "Result" in df.columns else ""
    }
    records.append(record)

# -----------------------------
# 6. เขียนออกเป็น JSONL
# -----------------------------
with open(output_path, "w", encoding="utf-8") as f:
    for r in records:
        f.write(json.dumps(r, ensure_ascii=False) + "\n")

print(f"\n✅ สร้างไฟล์ JSONL เรียบร้อย: {output_path}")
print(f"จำนวนแถวที่บันทึก: {len(records)}")
