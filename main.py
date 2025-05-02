import tempfile
import os
import openpyxl
from flask import Flask, request, jsonify, make_response
import urllib.parse
from flask_cors import CORS
from utils.excel_json_processor import format_shipping_document

# 創建 Flask 應用
app = Flask(__name__)
# 啟用 CORS
CORS(app, resources={r"/*": {"origins": "*"}})

@app.route('/', methods=['POST'])
@app.route('/process_excel', methods=['POST'])
def process_excel():
    """處理 Excel 文件或 JSON 資料的 API 端點
    
    可以接收：
    1. 上傳的 Excel 文件進行處理
    2. JSON 資料進行轉換和處理
    """
    try:
        # 檢查是否為 JSON 資料 (確認 Content-Type 是否包含 'application/json')
        if request.is_json:
            print("處理 JSON 資料...")
            # 獲取 JSON 資料
            json_data = request.get_json()
            
            # 創建臨時文件用於保存處理後的 Excel
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_output:
                temp_output_path = temp_output.name
            
            # 調用 format_shipping_document 函數處理 JSON 資料
            result_path = format_shipping_document(json_data, temp_output_path)
            
            if not result_path:
                return jsonify({"error": "處理 JSON 資料時出錯"}), 500
            
            # 讀取處理後的文件並創建響應
            with open(result_path, 'rb') as f:
                output_data = f.read()
            
            # 構建響應
            response = make_response(output_data)
            response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            
            # 設定文件名
            filename = "json_formatted_output.xlsx"
            
            # RFC 5987 格式編碼文件名
            encoded_filename = urllib.parse.quote(filename.encode('utf-8'))
            response.headers['Content-Disposition'] = f"attachment; filename*=UTF-8''{encoded_filename}"
            
            # 清理臨時文件
            try:
                os.unlink(temp_output_path)
            except Exception as e:
                print(f"清理臨時文件時出錯: {str(e)}")
            
            return response
            
        # 如果不是 JSON 資料，檢查是否有文件上傳
        elif 'file' in request.files:
            print("處理上傳的 Excel 文件...")
            uploaded_file = request.files['file']
            
            # 檢查文件名是否為空
            if uploaded_file.filename == '':
                return jsonify({"error": "文件名為空"}), 400
            
            # 檢查文件類型
            if not uploaded_file.filename.endswith(('.xlsx', '.xls')):
                return jsonify({"error": "只支持Excel文件(.xlsx或.xls)"}), 400
            
            # 創建臨時文件來保存上傳的Excel
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_input:
                temp_input_path = temp_input.name
                uploaded_file.save(temp_input_path)
            
            # 創建臨時文件用於保存處理後的Excel
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_output:
                temp_output_path = temp_output.name
            
            # 調用 format_shipping_document 函數處理文件
            result_path = format_shipping_document(temp_input_path, temp_output_path)
            
            if not result_path:
                return jsonify({"error": "處理Excel文件時出錯"}), 500
            
            # 讀取處理後的文件並創建響應
            with open(result_path, 'rb') as f:
                output_data = f.read()
            
            # 構建響應
            response = make_response(output_data)
            response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            
            # 處理文件名編碼問題 - 修復 Unicode 編碼錯誤
            filename = f"processed_{uploaded_file.filename}"
            
            # RFC 5987 格式編碼文件名
            encoded_filename = urllib.parse.quote(filename.encode('utf-8'))
            response.headers['Content-Disposition'] = f"attachment; filename*=UTF-8''{encoded_filename}"
            
            # 清理臨時文件
            try:
                os.unlink(temp_input_path)
                if result_path != temp_output_path:  # 如果結果路徑不是輸出臨時文件，則刪除臨時輸出文件
                    os.unlink(temp_output_path)
            except Exception as e:
                print(f"清理臨時文件時出錯: {str(e)}")
            
            return response
        else:
            return jsonify({"error": "未找到上傳的文件或JSON資料"}), 400
            
    except Exception as e:
        return jsonify({"error": f"處理請求時出錯: {str(e)}"}), 500

# 啟動應用
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=True)
