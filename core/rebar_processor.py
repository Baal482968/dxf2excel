"""
鋼筋處理相關功能模組
"""

import re
from config import REBAR_UNIT_WEIGHT, REBAR_DIAMETERS, REBAR_GRADES
from core.processors import get_processor, get_all_processors
# 圖形相關模組已移除，改為使用 assets/materials/ 資料夾中的圖示檔案

class RebarProcessor:
    """鋼筋處理器"""
    
    @staticmethod
    def extract_rebar_info(text):
        """
        從文字中提取鋼筋資訊
        
        格式範例：
        - #4@20
        - #4@20c/c
        - #4@20cm
        - #4@200mm
        """
        # 移除空白字元
        text = text.strip()
        
        # 基本格式檢查
        if not text or not text.startswith('#'):
            return None
        
        try:
            # 提取鋼筋編號
            rebar_number = re.match(r'#\d+', text).group()
            
            # 提取間距
            spacing_match = re.search(r'@(\d+)(?:c/c|cm|mm)?', text)
            if spacing_match:
                spacing = int(spacing_match.group(1))
                # 如果沒有單位，預設為公分
                if not re.search(r'@\d+(?:c/c|cm|mm)', text):
                    spacing *= 10  # 轉換為公釐
            else:
                spacing = None
            
            return {
                'rebar_number': rebar_number,
                'spacing': spacing,
                'diameter': RebarProcessor.get_rebar_diameter(rebar_number),
                'unit_weight': RebarProcessor.get_rebar_unit_weight(rebar_number),
                'grade': RebarProcessor.get_rebar_grade(rebar_number)
            }
        except Exception:
            return None

    @staticmethod
    def get_rebar_diameter(number):
        """獲取鋼筋直徑（公釐）"""
        return REBAR_DIAMETERS.get(number, 0)

    @staticmethod
    def get_rebar_unit_weight(number):
        """獲取鋼筋單位重量（kg/m）"""
        return REBAR_UNIT_WEIGHT.get(number, 0)

    @staticmethod
    def get_rebar_grade(number):
        """獲取鋼筋材質等級"""
        return REBAR_GRADES.get(number, "未知")

    @staticmethod
    def calculate_rebar_weight(number, length, count=1):
        """計算鋼筋重量（kg）"""
        unit_weight = RebarProcessor.get_rebar_unit_weight(number)
        return unit_weight * length * count

    @staticmethod
    def parse_rebar_text(text):
        """
        解析鋼筋文字格式 - 使用模組化處理器
        
        支援格式：
        - #3-700x99 (type10 單段直料)
        - 安#3-390x40 (type11 安全彎鉤直)
        - V113°#10-900+200x2 (type12 折料)
        - 弧450#10-700x1 (type18 直料圓弧)
        """
        text = text.strip()
        
        # 獲取所有處理器
        processors = get_all_processors()
        
        # 嘗試每個處理器
        for processor_type, processor in processors.items():
            if processor.can_process(text):
                print(f"🔍 使用 {processor_type} 處理器處理: {text}")
                result = processor.process(text)
                if result:
                    print(f"🔍 {processor_type} 處理結果: {result}")
                    return result
        
        # 無法解析的格式
        print(f"⚠️ 無法解析的鋼筋文字格式: {text}")
        return None

    @staticmethod
    def validate_rebar_number(number):
        """驗證鋼筋編號是否有效"""
        return number in REBAR_UNIT_WEIGHT

    @staticmethod
    def get_rebar_summary(rebar_list):
        """生成鋼筋統計摘要 - 使用基礎處理器的靜態方法"""
        from core.processors.base_processor import BaseRebarProcessor
        return BaseRebarProcessor.get_rebar_summary(rebar_list) 