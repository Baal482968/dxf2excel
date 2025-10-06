"""
Type10 直料 Excel 寫入器
"""

from .base_excel_writer import BaseExcelWriter


class Type10ExcelWriter(BaseExcelWriter):
    """Type10 直料 Excel 寫入器"""
    
    def __init__(self, graphics_manager=None):
        super().__init__(graphics_manager)
        self.rebar_type = 'type10'
    
    def get_rebar_type(self):
        """獲取鋼筋類型"""
        return self.rebar_type
    
    def generate_visual(self, rebar):
        """生成 Type10 鋼筋視覺表示"""
        segments = self._get_rebar_segments(rebar)
        rebar_id = rebar.get('raw_text', rebar.get('rebar_number', '#4'))
        
        # 檢查是否為 type10 鋼筋
        if self.graphics_available:
            try:
                # 生成 type10 鋼筋圖片
                length = segments[0] if segments else 0
                image = self.graphics_manager.generate_type10_rebar_image(length, rebar_id)
                
                if image:
                    temp_img_path = self._save_image_to_temp(image)
                    if temp_img_path:
                        print(f"🔍 生成 type10 鋼筋圖片: {temp_img_path}")
                        return temp_img_path
                        
            except Exception as e:
                print(f"⚠️ 生成 type10 鋼筋圖片失敗: {e}")
        
        # 如果圖片生成失敗，使用文字描述
        return self.generate_text_description(rebar)
