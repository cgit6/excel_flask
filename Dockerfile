FROM python:3.9-slim

# 設置工作目錄
WORKDIR /app

# 安裝依賴
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# 如果沒有 requirements.txt，可以取消註釋以下行並註釋上面兩行
# RUN pip install --no-cache-dir functions-framework flask openpyxl

# 複製應用程序文件
COPY . .

# 創建臨時目錄
RUN mkdir -p /tmp/excel_processing

# 設置環境變量
ENV PORT=8080

# 暴露端口
EXPOSE 8080

# 啟動命令
CMD exec functions-framework --target=process_excel --port=$PORT --debug 