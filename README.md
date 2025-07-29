# ERP 成本計算工具

一個專為 ERP 成本計算自動化設計的桌面應用程式，透過圖形化介面，讓使用者輕鬆匯入 Excel 檔案，並自動產生格式化、計算完成的成本分析報表。

## 主要功能

- 支援 Excel 檔案（.xlsx）匯入與自動驗證
- 自動產生格式化、帶有公式與樣式的成本計算報表
- 進度條與即時狀態提示
- 輸出檔案自動命名為 `ERP成本計算.xlsx`，可自訂輸出資料夾

## 安裝需求

1. **Python 版本**：建議 Python 3.8 以上
2. **必要套件**：請於專案根目錄執行以下指令安裝依賴
   ```bash
   pip install -r requirements.txt
   ```

## 使用說明

1. 執行主程式：
   ```bash
   python app/app.py
   ```

2. 打包成執行檔（可選）：
   ```bash
   pyinstaller --clean --onefile --noconsole --name "name" app/app.py
   ```
   打包完成後，執行檔會產生在 `dist` 資料夾中。

3. 操作步驟：
   - 點選「瀏覽」選擇輸入 Excel 檔案（需包含特定工作表）
   - 選擇輸出資料夾（預設為使用者下載資料夾）
   - 點擊「執行」開始處理
   - 處理完成後，會自動提示並開啟輸出檔案所在位置

## 輸入檔案格式要求

- 輸入 Excel 檔案需包含以下工作表（名稱需一致）：
  - 標準成本結構表
  - 鐵板重量計算
  - 鐵板材料費單價
  - 鐵板米數計算

## 技術架構

- **GUI 框架**：ttkbootstrap
- **Excel 處理**：openpyxl
- **檔案對話框/訊息提示**：tkinter.filedialog, tkinter.messagebox

## 主要檔案說明

- `app/app.py`：主視窗與操作流程
- `app/main.py`：成本計算主邏輯
- `app/excel.py`：Excel 內容處理與計算
- `app/style.py`：Excel 樣式與格式化輔助

## 注意事項

- 若輸出檔案已存在，會提示是否覆蓋
- 輸入檔案格式需正確，否則會顯示錯誤訊息
- 僅支援 .xlsx 格式

## 常見問題

- 若遇到「缺少工作表」或「格式錯誤」請確認輸入檔案內容
- 若無法啟動，請確認 Python 及相關套件已正確安裝