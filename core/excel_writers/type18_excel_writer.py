"""
Type18 直料圓弧 Excel 寫入器
"""

from .base_excel_writer import BaseExcelWriter


class Type18ExcelWriter(BaseExcelWriter):
    """Type18 直料圓弧 Excel 寫入器"""
    
    def __init__(self, graphics_manager=None):
        super().__init__(graphics_manager)
        self.rebar_type = 'type18'
    
    def get_rebar_type(self):
        """獲取鋼筋類型"""
        return self.rebar_type
    
    def generate_visual(self, rebar):
        """生成 Type18 鋼筋視覺表示"""
        segments = self._get_rebar_segments(rebar)
        radius = rebar.get('radius', 0)
        rebar_id = rebar.get('raw_text', rebar.get('rebar_number', '#4'))
        
        # 檢查是否為 type18 鋼筋
        if self.graphics_available:
            print(f"🔍 檢測到 type18 鋼筋，開始生成圖片...")
            try:
                # 生成 type18 鋼筋圖片
                length = segments[0] if segments else 0
                print(f"🔍 type18 長度: {length}, 半徑: {radius}, 號數: {rebar_id}")
                image = self.graphics_manager.generate_type18_rebar_image(length, radius, rebar_id)
                
                if image:
                    temp_img_path = self._save_image_to_temp(image)
                    if temp_img_path:
                        print(f"🔍 生成 type18 鋼筋圖片: {temp_img_path}")
                        return temp_img_path
                else:
                    print(f"⚠️ type18 圖片生成失敗，返回 None")
                    
            except Exception as e:
                print(f"⚠️ 生成 type18 鋼筋圖片失敗: {e}")
        else:
            print(f"⚠️ type18 檢測到但 graphics_available = {self.graphics_available}")
        
        # 如果圖片生成失敗，使用文字描述
        return self.generate_text_description(rebar)
