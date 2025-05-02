# Excel 文件處理 Cloud Function

這是一個 Google Cloud Functions 示例，用於接收上傳的 Excel 文件，處理後返回修改後的文件。

## 檔案結構

- `main.py` - Cloud Function 的主要代碼
- `requirements.txt` - 部署時需要的依賴包
- `test_upload.html` - 用於測試上傳功能的 HTML 頁面

## 本地測試

要在本地測試此 Cloud Function，請按照以下步驟操作：

1. 安裝依賴：

   ```bash
   pip install -r requirements.txt
   ```

2. 運行本地開發服務器：
   ```bash
   python demo.py
   ```
3. 服務器將在 http://localhost:8080 啟動

4. 使用瀏覽器打開 `test_upload.html` 文件，或直接使用 curl 測試：
   ```bash
   curl -X POST -F "file=@path/to/your/excel.xlsx" http://localhost:8080/process_excel --output processed.xlsx
   ```

## 部署到 Google Cloud Functions

### 使用 gcloud CLI 部署

1. 確保已安裝並配置 [Google Cloud SDK](https://cloud.google.com/sdk/docs/install)

2. 使用以下命令部署 Cloud Function：

   ```bash
   gcloud functions deploy process_excel \
     --runtime python310 \
     --trigger-http \
     --allow-unauthenticated \
     --entry-point process_excel \
     --region asia-east1
   ```

   > 注意：`--allow-unauthenticated` 允許公開訪問此 Cloud Function。在生產環境中，您可能需要添加適當的認證機制。

3. 部署完成後，您將收到一個 HTTPS URL，可以通過它訪問您的 Cloud Function。

### 使用 Google Cloud Console 部署

1. 前往 [Google Cloud Console](https://console.cloud.google.com/)
2. 開啟 Cloud Functions 頁面
3. 點擊「創建函數」
4. 填寫基本信息（名稱、地區等）
5. 設置觸發條件為 HTTP
6. 設置運行時為 Python 3.10
7. 在「入口點」字段中填入 `process_excel`
8. 上傳源代碼或使用內聯編輯器粘貼代碼
9. 設置 `requirements.txt`
10. 點擊「部署」

## 更新 test_upload.html

部署完成後，請編輯 `test_upload.html` 文件中的 `apiUrl` 變量，將其設置為您的 Cloud Function URL：

```javascript
const apiUrl =
  "https://your-region-your-project.cloudfunctions.net/process_excel";
```

然後在瀏覽器中打開此 HTML 文件以測試您的 Cloud Function。

## 自定義 Excel 處理邏輯

如需自定義 Excel 處理邏輯，請修改 `demo.py` 中的 `process_excel_file` 函數。目前，此函數執行以下操作：

1. 將第一行設置為粗體
2. 如果 A1 單元格為空，則添加標題
3. 將所有單元格中的文本轉換為大寫

您可以根據需要修改此邏輯，或者集成現有的 Excel 處理函數（如 test3.py 中的函數）。

測試用的命令行

```cmd
curl -X POST -F "file=@demo/翰禹_客戶出貨單_2025-04-15T11_30_55.xlsx" https://excel-flask-755089340805.us-central1.run.app/ --output output/processed_result.xlsx
```
