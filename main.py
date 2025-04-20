"""
運行Excel處理應用的主入口點
此文件僅用於Docker容器的入口點
"""
from demo import process_excel
import os
from flask import Flask, request

app = Flask(__name__)

@app.route('/', methods=['GET'])
def index():
    return process_excel(request)

@app.route('/process_excel', methods=['GET', 'POST'])
def excel_process_endpoint():
    return process_excel(request)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port) 