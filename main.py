import tempfile
import os
import openpyxl
from flask import Flask, jsonify, make_response, render_template_string
import functions_framework
import urllib.parse

@functions_framework.http
def process_excel_http(request):
    """Cloud Functions 入口點函數，處理 Excel 文件
    
    Args:
        request (flask.Request): HTTP 請求對象
        
    Returns:
        適當的 HTTP 響應
    """
    # 如果是GET請求，則顯示上傳表單
    if request.method == 'GET':
        html = '''
        <!DOCTYPE html>
        <html>
        <head>
            <title>Excel文件處理</title>
            <style>
                body { font-family: Arial, sans-serif; margin: 0; padding: 20px; line-height: 1.6; }
                .container { max-width: 800px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 5px; }
                h1 { color: #333; }
                .form-group { margin-bottom: 15px; }
                label { display: block; margin-bottom: 5px; font-weight: bold; }
                .btn { background-color: #4CAF50; color: white; padding: 10px 15px; border: none; cursor: pointer; }
                .btn:hover { background-color: #45a049; }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>Excel文件處理</h1>
                <p>請選擇要上傳的Excel文件 (.xlsx 或 .xls)</p>
                
                <form method="post" enctype="multipart/form-data">
                    <div class="form-group">
                        <label for="file">選擇文件:</label>
                        <input type="file" id="file" name="file" accept=".xlsx,.xls" required>
                    </div>
                    <button type="submit" class="btn">上傳並處理</button>
                </form>
            </div>
        </body>
        </html>
        '''
        return render_template_string(html)
    
    # 檢查請求方法
    if request.method != 'POST':
        return jsonify({"error": "只支持POST請求"}), 405
    
    # 檢查是否有文件上傳
    if 'file' not in request.files:
        return jsonify({"error": "未找到上傳的文件"}), 400
    
    uploaded_file = request.files['file']
    
    # 檢查文件名是否為空
    if uploaded_file.filename == '':
        return jsonify({"error": "文件名為空"}), 400
    
    # 檢查文件類型
    if not uploaded_file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({"error": "只支持Excel文件(.xlsx或.xls)"}), 400
    
    try:
        # 創建臨時文件來保存上傳的Excel
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_input:
            temp_input_path = temp_input.name
            uploaded_file.save(temp_input_path)
        
        # 創建臨時文件用於保存處理後的Excel
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_output:
            temp_output_path = temp_output.name
        
        # 處理Excel文件
        process_result = process_excel_file(temp_input_path, temp_output_path)
        
        if not process_result:
            return jsonify({"error": "處理Excel文件時出錯"}), 500
        
        # 讀取處理後的文件並創建響應
        with open(temp_output_path, 'rb') as f:
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
            os.unlink(temp_output_path)
        except Exception as e:
            print(f"清理臨時文件時出錯: {str(e)}")
        
        return response
        
    except Exception as e:
        return jsonify({"error": f"處理請求時出錯: {str(e)}"}), 500

def process_excel_file(input_path, output_path):
    """處理Excel文件的函數
    
    Args:
        input_path: 輸入Excel文件路徑
        output_path: 輸出Excel文件路徑
        
    Returns:
        bool: 處理是否成功
    """
    try:
        # 加載工作簿
        wb = openpyxl.load_workbook(input_path)
        ws = wb.active
        
        # 在這裡添加您的Excel處理邏輯
        # 例如：為第一行添加粗體
        for cell in ws[1]:
            cell.font = openpyxl.styles.Font(bold=True)
            
        # 示例：在A1單元格添加標題
        ws['A1'] = ws['A1'].value or "處理後的數據"
        
        # 示例：修改所有單元格中的文本
        for row in ws.iter_rows(min_row=2):  # 從第二行開始
            for cell in row:
                if isinstance(cell.value, str):
                    cell.value = cell.value.upper()  # 將文本轉換為大寫
        
        # 保存處理後的工作簿
        wb.save(output_path)
        return True
        
    except Exception as e:
        print(f"處理Excel文件時出錯: {str(e)}")
        return False 