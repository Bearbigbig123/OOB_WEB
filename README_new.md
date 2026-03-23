# OSAT SPC System

統計製程管制 (SPC) 分析系統，提供網頁介面與 API 服務，支援 OOB 分析、Tool Matching 分析、和 SPC CPK Dashboard。

## 系統架構

- **前端**：Streamlit 網頁應用程式 (`streamlit_app.py`)
- **後端**：FastAPI REST API 服務 (`main.py`)
- **分析引擎**：原有的 SPC 分析邏輯 (`oob_eng.py`, `tool_matching_widget_osat.py`)

## 快速啟動

### 1. 環境準備

建議使用虛擬環境：

```cmd
python -m venv .venv
.venv\Scripts\activate
```

安裝相依套件：

```cmd
pip install -r requirements.txt
```

### 2. 啟動後端 API 服務

在命令提示字元中執行：

```cmd
uvicorn main:app --host localhost --port 8000 --reload
```

成功啟動後可訪問：
- API 文件：http://localhost:8000/docs
- 健康檢查：http://localhost:8000/health

### 3. 啟動前端網頁介面

開啟另一個命令提示字元，執行：

```cmd
streamlit run app.py --server.runOnSave true
```

或者使用預設埠口：

```cmd
streamlit run streamlit_app.py
```

成功啟動後會自動開啟瀏覽器，或手動訪問：http://localhost:8501

### 4. 登入系統

系統預設帳號密碼：
- 帳號：`admin`
- 密碼：`password`

可透過環境變數自訂：
```cmd
set OOB_USER=your_username
set OOB_PASS=your_password
```

## 功能模組

### Split Chart
將大型 CSV 檔案依據 Chart 資訊分割成個別檔案，支援：
- Type2 垂直分割
- Type3 水平分割

### OOB/SPC 分析
執行統計製程管制分析，產生：
- SPC 控制圖
- 週報控制圖
- 違規規則檢測
- Excel 報告輸出

### Tool Matching
工具匹配分析，包含：
- Mean/Sigma 指標分析
- 統計檢定
- 分組比較

### SPC CPK Dashboard
製程能力分析儀表板：
- 多時間窗口 CPK 計算
- 趨勢分析 (R1/R2 衰退率)
- K 值 (偏移度) 計算
- 互動式圖表與統計摘要

## 檔案結構

```
├── main.py                          # FastAPI 後端服務
├── streamlit_app.py                 # Streamlit 前端應用
├── oob_eng.py                       # SPC 分析核心邏輯
├── tool_matching_widget_osat.py     # Tool Matching 分析
├── spc_cpk_dashboard_osat.py        # CPK Dashboard (PyQt 版本)
├── requirements.txt                 # Python 相依套件
├── input/                           # 輸入資料夾
│   ├── All_Chart_Information.xlsx   # Chart 設定檔
│   └── raw_charts/                  # 原始資料 CSV 檔案
├── output/                          # 輸出圖表資料夾
└── temp_uploads/                    # 暫存上傳檔案
```

## API 端點

- `GET /health` - 服務健康檢查
- `GET /` - API 資訊與預設路徑
- `POST /process` - OOB/SPC 分析
- `POST /split` - CSV 檔案分割
- `POST /tool-matching` - Tool Matching 分析
- `POST /spc-cpk` - SPC CPK Dashboard 分析

## 注意事項

- 確保兩個服務都在運行才能正常使用網頁介面
- 圖表生成使用 matplotlib Agg 後端，適合無頭伺服器環境
- 支援 Excel 檔案上傳與下載
- 所有分析結果會暫存在記憶體中，重啟服務會清除

## 故障排除

### 常見問題

1. **API 連線失敗**
   - 確認後端服務正在運行：`http://localhost:8000/health`
   - 檢查防火牆設定
   - 確認埠口 8000 未被其他服務佔用

2. **登入失敗**
   - 使用預設帳密：admin / password
   - 檢查環境變數設定

3. **檔案上傳問題**
   - 確認檔案格式正確 (Excel 或 CSV)
   - 檢查檔案大小限制
   - 確認 temp_uploads 資料夾有寫入權限

4. **圖表顯示異常**
   - 檢查 output 資料夾權限
   - 確認原始資料格式正確
   - 查看後端 console 錯誤訊息

### 開發模式

如需修改程式碼，建議使用以下命令啟動開發模式：

```cmd
# 後端 (自動重載)
uvicorn main:app --reload --host localhost --port 8000

# 前端 (檔案變更自動重載)
streamlit run streamlit_app.py --server.runOnSave true
```