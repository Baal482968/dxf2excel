"""
Type19 直段+弧段 Excel 寫入器
"""

from .base_excel_writer import BaseExcelWriter


class Type19ExcelWriter(BaseExcelWriter):
    """Type19 直段+弧段 Excel 寫入器"""
    
    def __init__(self, graphics_manager=None):
        super().__init__(graphics_manager)
        self.rebar_type = 'type19'
    
    def get_rebar_type(self):
        """獲取鋼筋類型"""
        return self.rebar_type
    
    def generate_visual(self, rebar):
        """生成 Type19 鋼筋視覺表示"""
        segments = self._get_rebar_segments(rebar)
        radius = rebar.get('radius', 0)
        rebar_id = rebar.get('raw_text', rebar.get('rebar_number', '#4'))
        
        # 檢查是否為 type19 鋼筋
        if self.graphics_available:
            print(f"🔍 檢測到 type19 鋼筋，開始生成圖片...")
            try:
                # 生成 type19 鋼筋圖片
                straight_length = segments[0] if len(segments) > 0 else 0
                arc_length = segments[1] if len(segments) > 1 else 0
                print(f"🔍 type19 直段: {straight_length}, 弧段: {arc_length}, 半徑: {radius}, 號數: {rebar_id}")
                image = self.graphics_manager.generate_type19_rebar_image(straight_length, arc_length, radius, rebar_id)
                
                if image:
                    temp_img_path = self._save_image_to_temp(image)
                    if temp_img_path:
                        print(f"🔍 生成 type19 鋼筋圖片: {temp_img_path}")
                        return temp_img_path
                else:
                    print(f"⚠️ type19 圖片生成失敗，返回 None")
                    
            except Exception as e:
                print(f"⚠️ 生成 type19 鋼筋圖片失敗: {e}")
        else:
            print(f"⚠️ type19 檢測到但 graphics_available = {self.graphics_available}")
        
        # 如果圖片生成失敗，使用文字描述
        return self.generate_text_description(rebar)
