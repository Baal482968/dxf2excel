"""
鋼筋處理器 - 重構版
使用模組化設計，支援多種鋼筋類型
"""

from .processors import get_processor, get_all_processors

class RebarProcessor:
    """鋼筋處理器 - 主協調器"""
    
    def __init__(self):
        self.processors = get_all_processors()
    
    def parse_rebar_text(self, text):
        """
        解析鋼筋文字格式
        
        支援格式：
        - #3-700x99 (type10 單段直料)
        - 安#3-390x40 (type11 安全彎鉤直)
        - V113°#10-900+200x2 (type12 折料)
        """
        text = text.strip()
        
        # 嘗試所有處理器
        for processor in self.processors.values():
            if processor.can_process(text):
                result = processor.process(text)
                if result:
                    print(f"🔍 使用 {processor.rebar_type} 處理器解析: {text}")
                    return result
        
        # 無法解析的格式
        print(f"⚠️ 無法解析鋼筋文字: {text}")
        return None
    
    def get_supported_types(self):
        """獲取支援的鋼筋類型"""
        return list(self.processors.keys())
    
    def add_processor(self, rebar_type, processor):
        """添加新的處理器"""
        self.processors[rebar_type] = processor
    
    def remove_processor(self, rebar_type):
        """移除處理器"""
        if rebar_type in self.processors:
            del self.processors[rebar_type]
    
    # 保留舊的靜態方法以維持向後相容性
    @staticmethod
    def extract_rebar_info(text):
        """從文字中提取鋼筋資訊（向後相容）"""
        # 這裡可以保留舊的邏輯或重定向到新方法
        pass
    
    @staticmethod
    def get_rebar_diameter(number):
        """獲取鋼筋直徑（公釐）"""
        from config import REBAR_DIAMETERS
        return REBAR_DIAMETERS.get(number, 0)
    
    @staticmethod
    def get_rebar_unit_weight(number):
        """獲取鋼筋單位重量（kg/m）"""
        from config import REBAR_UNIT_WEIGHT
        return REBAR_UNIT_WEIGHT.get(number, 0)
    
    @staticmethod
    def get_rebar_grade(number):
        """獲取鋼筋材質等級"""
        from config import REBAR_GRADES
        return REBAR_GRADES.get(number, "未知")
    
    @staticmethod
    def calculate_rebar_weight(number, length, count=1):
        """計算鋼筋重量（kg）"""
        from config import REBAR_UNIT_WEIGHT
        unit_weight = REBAR_UNIT_WEIGHT.get(number, 0)
        return unit_weight * length * count
    
    @staticmethod
    def validate_rebar_number(number):
        """驗證鋼筋編號是否有效"""
        from config import REBAR_UNIT_WEIGHT
        return number in REBAR_UNIT_WEIGHT
    
    @staticmethod
    def get_rebar_summary(rebar_list):
        """生成鋼筋統計摘要"""
        summary = {}
        
        for rebar in rebar_list:
            number = rebar['rebar_number']
            if number not in summary:
                summary[number] = {
                    'count': 0,
                    'total_length': 0,
                    'total_weight': 0,
                    'diameter': RebarProcessor.get_rebar_diameter(number),
                    'grade': RebarProcessor.get_rebar_grade(number)
                }
            
            summary[number]['count'] += 1
            summary[number]['total_length'] += rebar.get('length', 0)
            summary[number]['total_weight'] += rebar.get('weight', 0)
        
        return summary
