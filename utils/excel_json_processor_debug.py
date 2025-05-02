import os
import sys
import json
from utils.excel_json_processor import format_shipping_document

# 指定測試用的JSON資料
json_data = '''[
    {
        "name": "黃萍",
        "products": [
            {"productName": "購物金", "amount": "1"},
            {"productName": "預購(過膝靴/長靴)歐規尺碼 36-42 0119（樣式：37）", "amount": "1"}
        ]
    },
    {
        "name": "廖素津",
        "products": [
            {"productName": "購物金", "amount": "1"},
            {"productName": "預購(過膝靴/長靴)歐規尺碼 36-42 0119（樣式：37）", "amount": "1"}
        ]
    },
    {
        "name": "李淑娟",
        "products": [
            {"productName": "(預購)-YJ00001-潮流內刷毛外套(灰/橘/白)（樣式：橘L）", "amount": "1"}
        ]
    }
]'''

# 測試JSON轉換並保存中間結果
def test_json_conversion_with_debug():
    print("======== 測試JSON資料輸入與中間結果輸出 ========")
    output_file = os.path.join('output', 'debug_output.xlsx')
    
    # 執行轉換，並保存中間處理結果
    result = format_shipping_document(json_data, output_file, export_pdf=False, save_intermediate=True)
    print(f"處理結果: {'成功' if result else '失敗'}")
    
    return result

if __name__ == "__main__":
    test_json_conversion_with_debug() 