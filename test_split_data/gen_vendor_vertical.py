"""
生成 Vendor_Vertical 模式的測試資料
--------------------------------------
欄位格式：
  Part ID | Item Name | Report Time | Lot Mean | Vendor Site | (其他附加欄位)

執行後產生：
  test_Vendor_Vertical.csv          <- 上傳給 Split 功能使用
  All_Chart_Information_vendor.xlsx <- 配套的 AllChart 控制線檔案
"""

import pandas as pd
import numpy as np
import os

np.random.seed(77)
os.makedirs(os.path.dirname(__file__) or ".", exist_ok=True)
OUT_DIR = os.path.dirname(os.path.abspath(__file__))

# ============================================================
# 參數設定
# ============================================================
PARTS   = ["PART-A01", "PART-A02", "PART-B01"]
ITEMS   = ["Thickness", "Roughness", "Hardness", "Flatness"]
VENDORS = ["VendorTW", "VendorJP", "VendorKR"]
N_LOTS  = 60   # 每個 Vendor 每個品項的批次數

# 各品項的製程目標值與標準差
ITEM_PARAMS = {
    "Thickness": dict(target=120.0, sigma=3.0,  ucl=129.0, lcl=111.0, usl=132.0, lsl=108.0),
    "Roughness": dict(target=0.80,  sigma=0.05, ucl=0.95,  lcl=0.65,  usl=1.00,  lsl=0.60),
    "Hardness":  dict(target=55.0,  sigma=2.0,  ucl=61.0,  lcl=49.0,  usl=63.0,  lsl=47.0),
    "Flatness":  dict(target=0.20,  sigma=0.02, ucl=0.26,  lcl=0.14,  usl=0.28,  lsl=0.12),
}

# ============================================================
# 1. 生成 CSV
# ============================================================
records = []
base_time = pd.Timestamp("2025-01-01")

for part in PARTS:
    for item in ITEMS:
        p = ITEM_PARAMS[item]
        for vendor in VENDORS:
            for i in range(N_LOTS):
                # 不同 Vendor 略微偏移，製造出可辨識的差異
                v_offset = (VENDORS.index(vendor) - 1) * p["sigma"] * 0.4
                val = round(np.random.normal(p["target"] + v_offset, p["sigma"]), 4)
                records.append({
                    "Part ID":     part,
                    "Item Name":   item,
                    "Report Time": (base_time + pd.Timedelta(hours=6 * i)).strftime("%Y/%m/%d %H:%M"),
                    "Lot Mean":    val,
                    "Vendor Site": vendor,
                    "Lot No":      f"L{str(i + 1).zfill(3)}",
                    "Inspector":   f"ENG{np.random.randint(1, 4)}",
                    "Equipment":   f"EQP-{np.random.randint(1, 5):02d}",
                })

df_csv = pd.DataFrame(records)
csv_path = os.path.join(OUT_DIR, "test_Vendor_Vertical.csv")
df_csv.to_csv(csv_path, index=False, encoding="utf-8-sig")
print(f"[OK] CSV  -> {csv_path}  ({len(df_csv)} rows)")

# ============================================================
# 2. 生成 All_Chart_Information（拆分後 GroupName=Part ID, ChartName=Item Name）
# ============================================================
allchart_rows = []
for i_p, part in enumerate(PARTS):
    for i_c, item in enumerate(ITEMS):
        p = ITEM_PARAMS[item]
        allchart_rows.append({
            "GroupName":      part,
            "ChartName":      item,
            "ChartID":        f"VND{str(i_p * len(ITEMS) + i_c + 1).zfill(3)}",
            "Material_no":    f"MAT_VND_{part}_{item[:3].upper()}",
            "Target":         p["target"],
            "UCL":            p["ucl"],
            "LCL":            p["lcl"],
            "USL":            p["usl"],
            "LSL":            p["lsl"],
            "Characteristics":"Nominal",
            "DetectionLimit": None,
            "ExpectedPattern":"Normal",
            "SampleCount":    N_LOTS * len(VENDORS),
            "Resolution":     0.0001 if p["sigma"] < 0.1 else 0.01,
        })

df_allchart = pd.DataFrame(allchart_rows)
xlsx_path = os.path.join(OUT_DIR, "All_Chart_Information_vendor.xlsx")
df_allchart.to_excel(xlsx_path, index=False)
print(f"[OK] XLSX -> {xlsx_path}  ({len(df_allchart)} charts)")
print()
print(df_allchart[["GroupName", "ChartName", "Target", "UCL", "LCL"]].to_string(index=False))
