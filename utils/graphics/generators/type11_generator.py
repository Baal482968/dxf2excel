"""
Type11 安全彎鉤直圖形生成器
"""

from .base_generator import BaseImageGenerator

class Type11ImageGenerator(BaseImageGenerator):
    """Type11 安全彎鉤直圖形生成器"""
    
    def __init__(self):
        super().__init__()
        self.rebar_type = 'type11'
    
    def get_material_prefix(self):
        """獲取材料目錄前綴"""
        return "11-"
    
    def generate_image(self, length, rebar_number, available_materials):
        """生成 type11 鋼筋圖片"""
        try:
            # 尋找 type11 材料
            type11_material = self.find_material(available_materials)
            if not type11_material:
                print(f"❌ 找不到 type11 材料")
                return None
            
            # 構建 SVG 檔案路徑
            svg_path = self.get_svg_path(type11_material)
            if not svg_path.exists():
                print(f"❌ SVG 檔案不存在: {svg_path}")
                return None
            
            # 解析 SVG 並生成圖片
            return self._create_image_from_svg(svg_path, length, rebar_number)
            
        except Exception as e:
            print(f"❌ 生成 type11 鋼筋圖片失敗: {e}")
            return None
    
    def _create_image_from_svg(self, svg_path, length, rebar_number):
        """從 SVG 創建 type11 鋼筋圖片"""
        try:
            # 解析 SVG
            root = self.parse_svg(svg_path)
            if root is None:
                return None
            
            # 創建圖片
            img_width = 800
            img_height = 400
            image, draw = self.create_base_image(img_width, img_height)
            
            # 解析 SVG 中的 path 數據
            path_element = root.find(".//{http://www.w3.org/2000/svg}path")
            if path_element is not None:
                # 計算縮放比例
                scale_x = img_width / 800
                scale_y = img_height / 600
                
                # 繪製鋼筋線條
                line_width = 8
                
                # 1. 主要水平線段
                x1 = int(50 * scale_x)
                y1 = int(336.89 * scale_y)
                x2 = int(750 * scale_x)
                y2 = int(336.89 * scale_y)
                draw.line([(x1, y1), (x2, y2)], fill='black', width=line_width)
                
                # 2. 垂直彎折線段
                x3 = int(750 * scale_x)
                y3 = int(336.89 * scale_y)
                x4 = int(750 * scale_x)
                y4 = int(263.11 * scale_y)
                draw.line([(x3, y3), (x4, y4)], fill='black', width=line_width)
                
                # 3. 短水平線段
                x5 = int(750 * scale_x)
                y5 = int(263.11 * scale_y)
                x6 = int(661.46 * scale_x)
                y6 = int(263.11 * scale_y)
                draw.line([(x5, y5), (x6, y6)], fill='black', width=line_width)
                
                # 添加長度標註
                font = self.get_font(36)
                length_text = str(int(length))
                text_x = (x1 + x2) // 2
                text_y = y1 - 50
                self.draw_text_centered(draw, length_text, text_x, text_y, font)
                
            else:
                # 使用預設繪製
                self._draw_default_shape(draw, length, img_width, img_height)
            
            return image
            
        except Exception as e:
            print(f"❌ 從 SVG 創建 type11 圖片失敗: {e}")
            return None
    
    def _draw_default_shape(self, draw, length, img_width, img_height):
        """繪製預設形狀"""
        padding = 100
        line_start_x = padding
        line_end_x = img_width - padding
        line_y = img_height // 2
        
        # 繪製主要直線段
        draw.line([(line_start_x, line_y), (line_end_x - 100, line_y)], fill='black', width=8)
        
        # 繪製彎鉤
        hook_start_x = line_end_x - 100
        hook_end_x = line_end_x - 50
        hook_height = 50
        
        draw.line([(hook_start_x, line_y), (hook_start_x, line_y - hook_height)], fill='black', width=8)
        draw.line([(hook_start_x, line_y - hook_height), (hook_end_x, line_y - hook_height)], fill='black', width=8)
        
        # 添加長度標註
        font = self.get_font(36)
        length_text = str(int(length))
        text_x = (line_start_x + line_end_x - 100) // 2
        text_y = line_y - 50
        self.draw_text_centered(draw, length_text, text_x, text_y, font)
