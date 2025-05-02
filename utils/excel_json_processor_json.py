import os
import sys
import json
from utils.excel_json_processor import format_shipping_document

# 指定測試用的JSON資料
json_data = '''[
            {
                "name": "黃萍",
                "products": [
                    {
                        "productName": "購物金",
                        "amount": "1"
                    },
                    {
                        "productName": "預購(過膝靴/長靴)歐規尺碼 36-42 0119（樣式：37）",
                        "amount": "1"
                    }
                ]
            },
            {
                "name": "廖素津",
                "products": [
                    {
                        "productName": "購物金",
                        "amount": "1"
                    },
                    {
                        "productName": "預購(過膝靴/長靴)歐規尺碼 36-42 0119（樣式：37）",
                        "amount": "1"
                    }
                ]
            },
            {
                "name": "李淑娟",
                "products": [
                    {
                        "productName": "(預購)-YJ00001-潮流內刷毛外套(灰/橘/白)（樣式：橘L）",
                        "amount": "1"
                    }
                ]
            }
        ]'''

# 測試JSON輸入的情境
def test_json_input():
    print("======== 測試JSON資料輸入 ========")
    output_file = os.path.join('output', 'json_test_output.xlsx')
    
    # 嘗試使用字串形式的JSON資料
    print("\n----- 使用JSON字串輸入 -----")
    result1 = format_shipping_document(json_data, output_file, export_pdf=True)
    print(f"處理結果: {'成功' if result1 else '失敗'}")
    
    # 嘗試使用已解析的JSON物件
    print("\n----- 使用JSON物件輸入 -----")
    json_obj = json.loads(json_data)
    output_file2 = os.path.join('output', 'json_obj_test_output.xlsx')
    result2 = format_shipping_document(json_obj, output_file2, export_pdf=True)
    print(f"處理結果: {'成功' if result2 else '失敗'}")
    
    # 測試錯誤情況 - 缺少必要欄位
    print("\n----- 測試錯誤情況：缺少必要欄位 -----")
    invalid_json = '[{"name":"測試客戶", "products":[{"productName":"測試商品"}]}]'  # 缺少amount欄位
    output_file3 = os.path.join('output', 'invalid_json_test_output.xlsx')
    result3 = format_shipping_document(invalid_json, output_file3, export_pdf=False)
    print(f"處理結果: {'成功' if result3 else '失敗'}")  # 應該會失敗
    
    return result1, result2, result3

if __name__ == "__main__":
    test_json_input() 