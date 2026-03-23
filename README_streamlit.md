# OOB & Tool Matching 分析系統使用說明

## 系統架構
本系統採用前後端分離的架構：
- **後端**: FastAPI (http://localhost:8000) - 提供 REST API 服務
- **前端**: Streamlit (http://localhost:8501) - 提供網頁使用者介面

## 快速開始

### 1. 環境準備
```bash
# 安裝必要套件
pip install -r requirements.txt
```

### 2. 啟動系統
#### 方法一：一鍵啟動 (Windows)
```cmd
start_system.bat
```

#### 方法二：手動啟動
```bash
# 啟動後端 API 服務
uvicorn main:app --host localhost --port 8000 --reload

# 另開終端，啟動前端界面
streamlit run streamlit_app.py --server.port 8501
```

### 3. 使用網頁界面
開啟瀏覽器前往: http://localhost:8501

## 功能說明

### OOB/SPC 分析分頁
- **檔案上傳**: 支援上傳 Chart Information Excel 檔案和原始 CSV 資料
- **參數設定**: 
  - 儲存 Excel 報告選項
  - 圖片縮放比例調整
  - 限制處理圖表數量
- **結果展示**: 分析摘要、詳細結果表格、CSV 下載

### Tool Matching 分析分頁
- **檔案上傳**: 支援上傳 Tool Matching CSV 檔案
- **分析方法**: 
  - 指標分析 (Mean Index & Sigma Index)
  - 統計檢定 (paired/unpaired t-test)
- **參數設定**: 閾值、樣本數、資料過濾模式等
- **結果展示**: 匹配率統計、詳細結果表格、CSV 下載

## API 端點

### OOB 分析
- **POST** `/process` - 執行 OOB/SPC 分析
- **GET** `/health` - 檢查後台服務狀態

### Tool Matching 分析
- **POST** `/tool-matching` - 執行 Tool Matching 分析

## 檔案結構
```
├── streamlit_app.py      # Streamlit 前端主程式
├── main.py               # FastAPI 後端主程式
├── oob_eng.py           # OOB 分析核心功能
├── tool_matching_widget_osat.py  # Tool Matching 核心功能
├── requirements.txt      # Python 套件依賴
├── start_system.bat     # Windows 一鍵啟動腳本
└── README.md            # 本說明文件
```

## 注意事項
1. 確保後端 API 服務正在運行，前端才能正常使用
2. 上傳的檔案會暫時儲存在 `temp_uploads` 目錄，處理完成後會自動清除
3. 大檔案處理可能需要較長時間，請耐心等待
4. 建議使用現代瀏覽器 (Chrome, Firefox, Edge) 以獲得最佳體驗

## 疑難排解
- 若 API 連線失敗，請檢查後端服務是否正常運行
- 若檔案上傳失敗，請檢查檔案格式是否正確
- 若分析結果異常，請檢查輸入資料的格式和完整性