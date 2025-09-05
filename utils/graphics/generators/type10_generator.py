"""
Type10 直料圖形生成器
"""

from .base_generator import BaseImageGenerator

class Type10ImageGenerator(BaseImageGenerator):
    """Type10 直料圖形生成器"""
    
    def __init__(self):
        super().__init__()
        self.rebar_type = 'type10'
    
    def get_material_prefix(self):
        """獲取材料目錄前綴"""
        return "10-"
    
    def generate_image(self, length, rebar_number, available_materials):
        """生成 type10 鋼筋圖片"""
        try:
            # 尋找 type10 材料
            type10_material = self.find_material(available_materials)
            if not type10_material:
                print(f"❌ 找不到 type10 材料")
                return None
            
            # 構建 SVG 檔案路徑
            svg_path = self.get_svg_path(type10_material)
            if not svg_path.exists():
                print(f"❌ SVG 檔案不存在: {svg_path}")
                return None
            
            # 解析 SVG 並生成圖片
            return self._create_image_from_svg(svg_path, length, rebar_number)
            
        except Exception as e:
            print(f"❌ 生成 type10 鋼筋圖片失敗: {e}")
            return None
    
    def _create_image_from_svg(self, svg_path, length, rebar_number):
        """從 SVG 創建 type10 鋼筋圖片"""
        try:
            # 解析 SVG
            root = self.parse_svg(svg_path)
            if root is None:
                return None
            
            # 尋找 line 元素
            line_element = root.find(".//{http://www.w3.org/2000/svg}line")
            if line_element is None:
                print(f"❌ SVG 中找不到 line 元素")
                return None
            
            # 創建圖片
            img_width = 1200
            img_height = 600
            image, draw = self.create_base_image(img_width, img_height)
            
            # 繪製鋼筋線條
            padding = 100
            line_start_x = padding
            line_end_x = img_width - padding
            line_y = img_height // 2
            
            draw.line([(line_start_x, line_y), (line_end_x, line_y)], fill='black', width=12)
            
            # 繪製長度文字
            text = str(int(length))
            font = self.get_font(72)
            text_x = (line_start_x + line_end_x) // 2
            text_y = line_y - 120
            self.draw_text_centered(draw, text, text_x, text_y, font)
            
            return image
            
        except Exception as e:
            print(f"❌ 從 SVG 創建 type10 圖片失敗: {e}")
            return None
