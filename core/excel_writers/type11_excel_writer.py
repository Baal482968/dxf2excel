"""
Type11 安全彎鉤直 Excel 寫入器
"""

from .base_excel_writer import BaseExcelWriter


class Type11ExcelWriter(BaseExcelWriter):
    """Type11 安全彎鉤直 Excel 寫入器"""
    
    def __init__(self, graphics_manager=None):
        super().__init__(graphics_manager)
        self.rebar_type = 'type11'
    
    def get_rebar_type(self):
        """獲取鋼筋類型"""
        return self.rebar_type
    
    def generate_visual(self, rebar):
        """生成 Type11 鋼筋視覺表示"""
        segments = self._get_rebar_segments(rebar)
        rebar_id = rebar.get('raw_text', rebar.get('rebar_number', '#4'))
        
        # 檢查是否為 type11 鋼筋
        if self.graphics_available:
            print(f"🔍 檢測到 type11 鋼筋，開始生成圖片...")
            try:
                # 生成 type11 鋼筋圖片
                length = segments[0] if segments else 0
                print(f"🔍 type11 長度: {length}, 號數: {rebar_id}")
                image = self.graphics_manager.generate_type11_rebar_image(length, rebar_id)
                
                if image:
                    temp_img_path = self._save_image_to_temp(image)
                    if temp_img_path:
                        print(f"🔍 生成 type11 鋼筋圖片: {temp_img_path}")
                        return temp_img_path
                else:
                    print(f"⚠️ type11 圖片生成失敗，返回 None")
                    
            except Exception as e:
                print(f"⚠️ 生成 type11 鋼筋圖片失敗: {e}")
        else:
            print(f"⚠️ type11 檢測到但 graphics_available = {self.graphics_available}")
        
        # 如果圖片生成失敗，使用文字描述
        return self.generate_text_description(rebar)
