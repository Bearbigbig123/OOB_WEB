"""
生成 Test_Horizontal 模式的測試資料
--------------------------------------
欄位格式：
  Part ID | FT Test End Time | Test Site | <Item1> | <Item2> | ... | <ItemN>

每一列代表「一個 Part 在一個 Test Site 的一次測試記錄」，
測試項目橫向展開為多個欄位（有部分欄位會有 NaN，模擬實際缺測情況）。

執行後產生：
  test_Test_Horizontal.csv              <- 上傳給 Split 功能使用
  All_Chart_Information_test_h.xlsx     <- 配套的 AllChart 控制線檔案
"""

import pandas as pd
import numpy as np
import os

np.random.seed(55)
OUT_DIR = os.path.dirname(os.path.abspath(__file__))

# ============================================================
# 參數設定
# ============================================================
PARTS      = ["DUT-001", "DUT-002", "DUT-003"]
TEST_SITES = ["SiteAlpha", "SiteBeta", "SiteGamma"]
N_RECORDS  = 60   # 每個 Part x Site 的測試筆數

# 測試項目與其製程規格
TEST_ITEMS = {
    "Vt_mV":      dict(target=450.0, sigma=15.0, ucl=495.0, lcl=405.0, usl=510.0, lsl=390.0),
    "Idsat_uA":   dict(target=800.0, sigma=30.0, ucl=890.0, lcl=710.0, usl=920.0, lsl=680.0),
    "Ioff_pA":    dict(target=5.0,   sigma=0.8,  ucl=7.4,   lcl=2.6,   usl=8.0,   lsl=2.0),
    "Gm_mS":      dict(target=12.0,  sigma=0.5,  ucl=13.5,  lcl=10.5,  usl=14.0,  lsl=10.0),
    "SS_mV_dec":  dict(target=68.0,  sigma=2.0,  ucl=74.0,  lcl=62.0,  usl=76.0,  lsl=60.0),
    "DIBL_mV_V":  dict(target=80.0,  sigma=5.0,  ucl=95.0,  lcl=65.0,  usl=100.0, lsl=60.0),
}

MISSING_RATE = 0.05   # 模擬 5% 的缺測比例

# ============================================================
# 1. 生成 CSV（橫向格式）
# ============================================================
records = []
base_time = pd.Timestamp("2025-03-01")

for part in PARTS:
    for site in TEST_SITES:
        # 不同 Site 有略微系統性偏移
        site_offset_factor = (TEST_SITES.index(site) - 1) * 0.3
        for i in range(N_RECORDS):
            row = {
                "Part ID":          part,
                "FT Test End Time": (base_time + pd.Timedelta(hours=4 * i)).strftime("%Y/%m/%d %H:%M"),
                "Test Site":        site,
            }
            for item, p in TEST_ITEMS.items():
                if np.random.rand() < MISSING_RATE:
                    row[item] = np.nan   # 模擬缺測
                else:
                    offset = site_offset_factor * p["sigma"] * 0.5
                    row[item] = round(np.random.normal(p["target"] + offset, p["sigma"]), 4)
            records.append(row)

df_csv = pd.DataFrame(records)
csv_path = os.path.join(OUT_DIR, "test_Test_Horizontal.csv")
df_csv.to_csv(csv_path, index=False, encoding="utf-8-sig")
print(f"[OK] CSV  -> {csv_path}  ({len(df_csv)} rows, {len(df_csv.columns)} columns)")
print(f"     欄位: {df_csv.columns.tolist()}")

# ============================================================
# 2. 生成 All_Chart_Information（拆分後 GroupName=Part ID, ChartName=測試項目欄名）
# ============================================================
allchart_rows = []
for i_p, part in enumerate(PARTS):
    for i_c, (item, p) in enumerate(TEST_ITEMS.items()):
        allchart_rows.append({
            "GroupName":      part,
            "ChartName":      item,
            "ChartID":        f"FT{str(i_p * len(TEST_ITEMS) + i_c + 1).zfill(3)}",
            "Material_no":    f"MAT_FT_{part}_{item[:4].upper()}",
            "Target":         p["target"],
            "UCL":            p["ucl"],
            "LCL":            p["lcl"],
            "USL":            p["usl"],
            "LSL":            p["lsl"],
            "Characteristics":"Nominal",
            "DetectionLimit": None,
            "ExpectedPattern":"Normal",
            "SampleCount":    N_RECORDS * len(TEST_SITES),
            "Resolution":     0.001 if p["sigma"] < 1 else 0.1,
        })

df_allchart = pd.DataFrame(allchart_rows)
xlsx_path = os.path.join(OUT_DIR, "All_Chart_Information_test_h.xlsx")
df_allchart.to_excel(xlsx_path, index=False)
print(f"[OK] XLSX -> {xlsx_path}  ({len(df_allchart)} charts)")
print()
print(df_allchart[["GroupName", "ChartName", "Target", "UCL", "LCL"]].to_string(index=False))
