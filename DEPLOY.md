# Excel 處理應用 - 部署指南

本文檔提供了使用 Docker 部署 Excel 處理應用的詳細步驟，包括本地部署和雲端部署。

## 文件結構
確保您的項目文件結構如下：
```
excel_flask/
├── demo.py          # 主要應用程序代碼
├── Dockerfile       # Docker 配置文件
├── requirements.txt # 依賴包列表
└── DEPLOY.md        # 本部署指南
```

## 本地 Docker 部署

### 1. 構建 Docker 映像
```bash
cd excel_flask
docker build -t excel-processor .
```

### 2. 運行 Docker 容器
```bash
docker run -p 8080:8080 excel-processor
```

### 3. 測試應用
應用將運行在 http://localhost:8080。您可以使用 cURL 或任何 API 測試工具發送請求：
```bash
curl -X POST -F "file=@path/to/your/excel.xlsx" http://localhost:8080/process_excel --output processed.xlsx
```

## 部署到 Google Cloud Run

### 1. 構建並推送到 Google Container Registry

首先，確保已安裝並設置好 [Google Cloud SDK](https://cloud.google.com/sdk/docs/install)。

```bash
# 設置項目 ID
PROJECT_ID=$(gcloud config get-value project)

# 構建映像
docker build -t gcr.io/$PROJECT_ID/excel-processor .

# 推送到 Google Container Registry
docker push gcr.io/$PROJECT_ID/excel-processor
```

### 2. 部署到 Cloud Run
```bash
gcloud run deploy excel-processor \
  --image gcr.io/$PROJECT_ID/excel-processor \
  --platform managed \
  --region asia-east1 \
  --allow-unauthenticated
```

部署完成後，您將收到服務 URL，可以通過該 URL 訪問您的應用。

## 部署到其他雲平台

### AWS App Runner
1. 將 Docker 映像推送到 Amazon ECR
2. 使用 AWS App Runner 服務部署該映像

### Azure Container Instances
1. 將 Docker 映像推送到 Azure Container Registry
2. 使用 Azure Container Instances 部署該映像

## 環境變量配置

您可以通過設置以下環境變量來配置應用：

- `PORT`：應用監聽的端口（默認為 8080）
- `DEBUG`：是否啟用調試模式（設置為 true 或 false）

例如，在 Google Cloud Run 中設置環境變量：
```bash
gcloud run deploy excel-processor \
  --image gcr.io/$PROJECT_ID/excel-processor \
  --set-env-vars="DEBUG=false"
```

## 故障排除

### 1. 構建錯誤
如果構建過程中出現錯誤，請檢查：
- requirements.txt 文件是否存在且格式正確
- 所有依賴包是否兼容

### 2. 啟動錯誤
如果容器啟動失敗，可以查看日誌：
```bash
docker logs <container_id>
```

### 3. 請求錯誤
如果 API 請求返回錯誤，請檢查：
- 請求格式是否正確
- 上傳的文件是否為有效的 Excel 文件 