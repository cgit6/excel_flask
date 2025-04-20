from flask import Flask, request, jsonify, make_response
import tempfile
import os
import openpyxl

app = Flask(__name__)

@app.route("/process_excel", methods=["POST"])
def process_excel():
    if request.method != 'POST':
        return jsonify({"error": "只支持POST请求"}), 405
    
    if 'file' not in request.files:
        return jsonify({"error": "未找到上传的文件"}), 400
    
    uploaded_file = request.files['file']
    
    if uploaded_file.filename == '':
        return jsonify({"error": "文件名为空"}), 400
    
    if not uploaded_file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({"error": "只支持Excel文件(.xlsx或.xls)"}), 400
    
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_input:
            temp_input_path = temp_input.name
            uploaded_file.save(temp_input_path)
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_output:
            temp_output_path = temp_output.name
        
        process_result = process_excel_file(temp_input_path, temp_output_path)
        
        if not process_result:
            return jsonify({"error": "处理Excel文件时出错"}), 500
        
        with open(temp_output_path, 'rb') as f:
            output_data = f.read()
        
        response = make_response(output_data)
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        response.headers['Content-Disposition'] = f'attachment; filename=processed_{uploaded_file.filename}'
        
        try:
            os.unlink(temp_input_path)
            os.unlink(temp_output_path)
        except Exception as e:
            print(f"清理临时文件时出错: {str(e)}")
        
        return response
        
    except Exception as e:
        return jsonify({"error": f"处理请求时出错: {str(e)}"}), 500

def process_excel_file(input_path, output_path):
    try:
        wb = openpyxl.load_workbook(input_path)
        ws = wb.active

        for cell in ws[1]:
            cell.font = openpyxl.styles.Font(bold=True)
        
        ws['A1'] = ws['A1'].value or "处理后的数据"
        
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, str):
                    cell.value = cell.value.upper()
        
        wb.save(output_path)
        return True
        
    except Exception as e:
        print(f"处理Excel文件时出错: {str(e)}")
        return False

# 啟動 Flask app
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
