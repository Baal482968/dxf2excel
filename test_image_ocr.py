#!/usr/bin/env python3
"""
測試圖片 OCR 功能
"""

import os
import sys
from core.cad_reader import CADReader

def test_image_ocr(dxf_file_path):
    """測試圖片 OCR 功能"""
    print(f"測試檔案: {dxf_file_path}")
    print("=" * 50)
    
    # 創建 CAD 讀取器
    reader = CADReader()
    
    # 開啟 DXF 檔案
    if not reader.open_file(dxf_file_path):
        print("❌ 無法開啟 DXF 檔案")
        return
    
    try:
        # 獲取圖面資訊
        info = reader.get_drawing_info()
        if info:
            print(f"📄 圖面資訊:")
            print(f"   - 檔案名稱: {info['filename']}")
            print(f"   - DXF 版本: {info['dxfversion']}")
            print(f"   - 圖層數量: {len(info['layers'])}")
            print(f"   - 實體數量: {info['entities_count']}")
            print()
        
        # 提取圖片中的數字
        print("🔍 開始提取圖片中的數字...")
        image_numbers = reader.extract_image_numbers()
        
        if image_numbers:
            print(f"✅ 成功提取到 {len(image_numbers)} 個數字:")
            for i, number_info in enumerate(image_numbers, 1):
                print(f"   {i}. 數字: {number_info['number_text']}")
                print(f"      信心度: {number_info['confidence']}%")
                print(f"      圖片路徑: {number_info['image_path']}")
                print(f"      插入點: {number_info['insert_point']}")
                print(f"      圖片尺寸: {number_info['image_size']}")
                print()
        else:
            print("❌ 沒有找到圖片實體或無法識別數字")
        
        # 提取文字實體（對比）
        print("🔍 提取文字實體...")
        rebar_texts = reader.extract_rebar_texts()
        print(f"✅ 找到 {len(rebar_texts)} 個文字實體")
        
    except Exception as e:
        print(f"❌ 測試過程中發生錯誤: {str(e)}")
    
    finally:
        # 關閉檔案
        reader.close_file()

def main():
    """主函數"""
    if len(sys.argv) != 2:
        print("使用方法: python3 test_image_ocr.py <DXF檔案路徑>")
        print("例如: python3 test_image_ocr.py assets/example1/example1.dxf")
        return
    
    dxf_file_path = sys.argv[1]
    
    if not os.path.exists(dxf_file_path):
        print(f"❌ 檔案不存在: {dxf_file_path}")
        return
    
    test_image_ocr(dxf_file_path)

if __name__ == "__main__":
    main() 