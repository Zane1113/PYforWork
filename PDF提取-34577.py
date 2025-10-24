import os
import re
import pandas as pd
import pdfplumber

# === 改成你自己的路径 ===
folder_path = "/Users/cenghongyi/Downloads/drive-download-20251009T030723Z-1-001"
output_file = os.path.join(folder_path, "invoice_data.xlsx")

# 正则
order_re = re.compile(r"\b[A-Z]{3}\d{6}\b")   # 订单号 ABC123456
longnum_re = re.compile(r"\b\d{6,}\b")        # >=6 位长数字（用于 invoice/customer）
amount_re = re.compile(r"\d[\d,\.]*")         # 金额匹配

rows = []

for fname in os.listdir(folder_path):
    if not fname.lower().endswith(".pdf"):
        continue
    path = os.path.join(folder_path, fname)
    invoice_num = None
    orders = []

    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            # 去掉空行但保留顺序
            lines = [ln.rstrip() for ln in text.splitlines() if ln.strip()]

            for i, line in enumerate(lines):
                up = line.upper()

                # ---- 找 Invoice（上下文） ----
                if invoice_num is None and "INVOICE" in up:
                    # 合并附近行作为 header context（上、当前、下）
                    context = " ".join(lines[max(0, i-1): i+2]).upper()
                    # 在当前或往下几行找包含长数字的那一行
                    for j in range(i, min(i+5, len(lines))):
                        nums = longnum_re.findall(lines[j])
                        if nums:
                            # 如果 header context 含 CUSTOMER，并且这一行有 >=2 个长数字 -> 取第二个
                            if "CUSTOMER" in context and len(nums) >= 2:
                                invoice_num = nums[1]
                            else:
                                # 否则取该行最后一个长数字（通常是最右的 invoice）
                                invoice_num = nums[-1]
                            break

                # ---- 找订单号 + 该行最右金额 ----
                for m in order_re.finditer(line):
                    order_no = m.group(0)
                    amounts = amount_re.findall(line)
                    amount = amounts[-1] if amounts else ""
                    orders.append((order_no, amount))

    # 把每个订单写到表里（发票号可能为空）
    for o, a in orders:
        rows.append({"Invoice Number": invoice_num, "Ref": o, "Amount": a})

# 导出
df = pd.DataFrame(rows, columns=["Invoice Number", "Ref", "Amount"])
df.to_excel(output_file, index=False)
print(f"完成 -> {output_file} , 共 {len(df)} 行")
