"""
基礎鋼筋處理器抽象類
"""

from abc import ABC, abstractmethod
from config import REBAR_UNIT_WEIGHT, REBAR_DIAMETERS, REBAR_GRADES

class BaseRebarProcessor(ABC):
    """鋼筋處理器基礎類"""
    
    def __init__(self):
        self.rebar_type = None
        self.pattern = None
    
    @abstractmethod
    def get_pattern(self):
        """獲取正則表達式模式"""
        pass
    
    @abstractmethod
    def parse_match(self, match, text):
        """解析匹配結果"""
        pass
    
    def can_process(self, text):
        """檢查是否能處理此文字"""
        import re
        pattern = self.get_pattern()
        return re.match(pattern, text) is not None
    
    def process(self, text):
        """處理鋼筋文字"""
        import re
        text = text.strip()
        pattern = self.get_pattern()
        match = re.match(pattern, text)
        
        if match:
            return self.parse_match(match, text)
        return None
    
    def get_rebar_diameter(self, number):
        """獲取鋼筋直徑（公釐）"""
        return REBAR_DIAMETERS.get(number, 0)
    
    def get_rebar_unit_weight(self, number):
        """獲取鋼筋單位重量（kg/m）"""
        return REBAR_UNIT_WEIGHT.get(number, 0)
    
    def get_rebar_grade(self, number):
        """獲取鋼筋材質等級"""
        return REBAR_GRADES.get(number, "未知")
    
    def calculate_weight(self, rebar_number, length, count=1):
        """計算鋼筋重量（kg）"""
        unit_weight = self.get_rebar_unit_weight(rebar_number)
        return unit_weight * length * count / 100  # 轉換為 kg
