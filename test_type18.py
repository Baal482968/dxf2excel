#!/usr/bin/env python3
"""
測試 Type18 直料圓弧鋼筋功能
"""

from core.rebar_processor_new import RebarProcessor
from utils.graphics.manager_new import GraphicsManager
from core.excel_writer_new import ExcelWriter

def test_type18_parsing():
    """測試 Type18 文字解析"""
    print("🧪 測試 Type18 文字解析")
    print("=" * 50)
    
    processor = RebarProcessor()
    
    # 測試文字
    test_texts = [
        "弧450#10-700x1",
        "弧300#8-500x2",
        "弧600#12-900x1"
    ]
    
    for text in test_texts:
        print(f"\n📝 測試文字: {text}")
        result = processor.parse_rebar_text(text)
        if result:
            print(f"✅ 解析成功:")
            print(f"   類型: {result['type']}")
            print(f"   號數: {result['rebar_number']}")
            print(f"   長度: {result['length']}")
            print(f"   半徑: {result['radius']}")
            print(f"   數量: {result['count']}")
            print(f"   重量: {result['weight']:.2f} kg")
            print(f"   備註: {result['note']}")
        else:
            print(f"❌ 解析失敗")

def test_type18_image_generation():
    """測試 Type18 圖片生成"""
    print("\n🧪 測試 Type18 圖片生成")
    print("=" * 50)
    
    graphics_manager = GraphicsManager()
    
    # 測試參數
    test_cases = [
        {"length": 700, "radius": 450, "rebar_number": "#10"},
        {"length": 500, "radius": 300, "rebar_number": "#8"},
        {"length": 900, "radius": 600, "rebar_number": "#12"}
    ]
    
    for i, case in enumerate(test_cases, 1):
        print(f"\n📸 測試案例 {i}: 長度={case['length']}, 半徑={case['radius']}, 號數={case['rebar_number']}")
        
        try:
            image = graphics_manager.generate_type18_rebar_image(
                case['length'], 
                case['radius'], 
                case['rebar_number']
            )
            
            if image:
                # 保存測試圖片
                output_path = f"test_type18_case_{i}.png"
                image.save(output_path)
                print(f"✅ 圖片生成成功: {output_path}")
            else:
                print(f"❌ 圖片生成失敗")
                
        except Exception as e:
            print(f"❌ 圖片生成錯誤: {e}")

def test_type18_excel_integration():
    """測試 Type18 Excel 整合"""
    print("\n🧪 測試 Type18 Excel 整合")
    print("=" * 50)
    
    # 創建測試資料
    test_data = [
        {
            'rebar_number': '#10',
            'segments': [700],
            'angles': [],
            'radius': 450,
            'count': 1,
            'raw_text': '弧450#10-700x1',
            'length': 700,
            'weight': 4.32,
            'type': 'type18',
            'note': '直料圓弧 R450'
        },
        {
            'rebar_number': '#8',
            'segments': [500],
            'angles': [],
            'radius': 300,
            'count': 2,
            'raw_text': '弧300#8-500x2',
            'length': 500,
            'weight': 3.95,
            'type': 'type18',
            'note': '直料圓弧 R300'
        }
    ]
    
    try:
        # 創建 Excel 寫入器
        writer = ExcelWriter(image_mode="mixed")
        writer.create_workbook()
        
        # 寫入標題
        header_row = writer.write_title("Type18 直料圓弧測試")
        
        # 寫入表頭
        writer.write_header()
        
        # 寫入資料
        next_row = writer.write_rebar_data(test_data, header_row + 1)
        
        # 寫入統計摘要
        summary_row = writer.write_summary(test_data, next_row)
        
        # 寫入頁尾
        writer.write_footer(summary_row + 1)
        
        # 格式化
        writer.format_worksheet()
        
        # 儲存
        output_path = "test_type18_output.xlsx"
        writer.save_workbook(output_path)
        
        print(f"✅ Excel 檔案生成成功: {output_path}")
        
    except Exception as e:
        print(f"❌ Excel 生成失敗: {e}")

def main():
    """主測試函數"""
    print("🚀 開始 Type18 直料圓弧鋼筋測試")
    print("=" * 60)
    
    # 測試文字解析
    test_type18_parsing()
    
    # 測試圖片生成
    test_type18_image_generation()
    
    # 測試 Excel 整合
    test_type18_excel_integration()
    
    print("\n🎉 Type18 測試完成！")

if __name__ == "__main__":
    main()
