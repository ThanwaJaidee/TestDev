import pandas as pd
import json
import os
import math


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


test_col = find_col(df, ["test"])
step_col = find_col(df, ["step", "no"])
desc_col = find_col(df, ["desc"])
unit_col = find_col(df, ["unit"])
setting_col = find_col(df, ["set"])
input_node_col = find_col(df, ["input node"])
input_value_col = find_col(df, ["value"])
meas_node_col = find_col(df, ["meas node"])
low_limit_col = find_col(df, ["low"])
high_limit_col = find_col(df, ["high"])
# result_col = find_col(df, ["result", "status", "output"])


print(f"🧾 Detected columns: {test_col} / {step_col} / {desc_col} / {unit_col} / {setting_col} / {input_node_col} / {input_value_col} / {meas_node_col} / {low_limit_col} / {high_limit_col}")



# ถ้าชื่อคอลัมน์ไม่ตรง สามารถเปลี่ยนชื่อได้เช่น
# df.rename(columns={"Step#": "Step #"}, inplace=True)

# -----------------------------
# 4. รวมข้อความเป็น instruction
# -----------------------------


# df["instruction"] = (
#     "Test "         + df[test_col].astype(str)  + " : " +
#     "Step "         + df[step_col].astype(str)  + " : " +
#     "Description "  + df[desc_col].fillna("")   + " : " +
#     "Unit "         + df[unit_col].fillna("")   + " : " 
# )

# ถ้าต้องการเพิ่มช่องอื่น เช่น “Low Limit” หรือ “High Limit” ก็ได้ เช่น
# df["instruction"] = (
#     "Step " + df["Step #"].astype(str) + ": " +
#     df["Description"].fillna("") +
#     " | Range: " + df["Low Limit"].astype(str) + " - " + df["High Limit"].astype(str)
# )
# ---------- helper สำหรับแปลงค่าให้เป็นชนิดที่ json รองรับ ----------
def to_jsonable(v):
    # แปลง NaN/None เป็น "" (หรือจะใช้ None ก็ได้ถ้าต้องการเก็บ null)
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    # แปลงชนิด numpy ให้เป็น Python พื้นฐาน
    if hasattr(v, "item"):
        try:
            return v.item()
        except Exception:
            pass
    return v
# -----------------------------
# 5. สร้างฟิลด์สำหรับ JSONL
# -----------------------------
records = []
for _, row in df.iterrows():
    result_obj = {
        "Test":        to_jsonable(row[test_col])       if test_col       else "",
        "Step":        to_jsonable(row[step_col])       if step_col       else "",
        "Description": to_jsonable(row[desc_col])       if desc_col       else "",
        "Unit":        to_jsonable(row[unit_col])       if unit_col       else "",
        # ใส่เพิ่มได้ถ้ามีคอลัมน์เหล่านี้
        "Setting":     to_jsonable(row[setting_col])    if setting_col    else "",
        "InputNode":   to_jsonable(row[input_node_col]) if input_node_col else "",
        "InputValue":  to_jsonable(row[input_value_col])if input_value_col else "",
        "MeasNode":    to_jsonable(row[meas_node_col])  if meas_node_col  else "",
        "LowLimit":    to_jsonable(row[low_limit_col])  if low_limit_col  else "",
        "HighLimit":   to_jsonable(row[high_limit_col]) if high_limit_col else "",
    }
    instruction_obj = " ".join(
    str(to_jsonable(val))
    for _, val in row.items()
    if pd.notna(val)
    )
    # instruction_obj = {
    #     to_jsonable(val)
    #     for val in row.items()
    #     if pd.notna(val)
    # }
    # instruction_obj = {
    #     str(col).strip(): to_jsonable(val)
    #     for col, val in row.items()
    #     if pd.notna(val)
    # }
    record = {
        # เก็บ instruction เป็น “สตริง JSON” เพื่อทำไฟล์ .jsonl ได้ง่าย
        "instruction": instruction_obj,
        # "instruction": json.dumps(instruction_obj, ensure_ascii=False),
        # "instruction": json.dumps(instruction_obj, ensure_ascii=False),
        # output จะเก็บผลลัพธ์/คำตอบจริงตอนทำ supervision ก็ได้
        "output": result_obj,
        # "output": json.dumps(result_obj, ensure_ascii=False),
        # "output": to_jsonable(row["Result"]) if "Result" in df.columns else ""
    }
    records.append(record)
print(instruction_obj)
# -----------------------------
# 6. เขียนออกเป็น JSONL
# -----------------------------
with open(output_path, "w", encoding="utf-8") as f:
    for r in records:
        f.write(json.dumps(r, ensure_ascii=False) + "\n")

print(f"\n✅ สร้างไฟล์ JSONL เรียบร้อย: {output_path}")
print(f"จำนวนแถวที่บันทึก: {len(records)}")
