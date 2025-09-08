"""
Type18 直料圓弧圖形生成器
"""

from .base_generator import BaseImageGenerator
import math

class Type18ImageGenerator(BaseImageGenerator):
    """Type18 直料圓弧圖形生成器"""
    
    def __init__(self):
        super().__init__()
        self.rebar_type = 'type18'
    
    def get_material_prefix(self):
        """獲取材料目錄前綴"""
        return "18-"
    
    def generate_image(self, length, radius, rebar_number, available_materials):
        """生成 type18 鋼筋圖片"""
        try:
            # 尋找 type18 材料
            type18_material = self.find_material(available_materials)
            if not type18_material:
                print(f"❌ 找不到 type18 材料")
                return None
            
            # 構建 SVG 檔案路徑
            svg_path = self.get_svg_path(type18_material)
            if not svg_path.exists():
                print(f"❌ SVG 檔案不存在: {svg_path}")
                return None
            
            # 解析 SVG 並生成圖片
            return self._create_image_from_svg(svg_path, length, radius, rebar_number)
            
        except Exception as e:
            print(f"❌ 生成 type18 鋼筋圖片失敗: {e}")
            return None
    
    def _create_image_from_svg(self, svg_path, length, radius, rebar_number):
        """從 SVG 創建 type18 鋼筋圖片"""
        try:
            # 解析 SVG
            root = self.parse_svg(svg_path)
            if root is None:
                return None
            
            # 創建圖片
            img_width = 800
            img_height = 400
            image, draw = self.create_base_image(img_width, img_height)
            
            # 解析 SVG 中的 path 數據（圓弧）
            path_element = root.find(".//{http://www.w3.org/2000/svg}path")
            if path_element is not None:
                # 計算縮放比例
                scale_x = img_width / 800
                scale_y = img_height / 600
                
                # 繪製圓弧鋼筋
                self._draw_arc_rebar(draw, length, radius, img_width, img_height, scale_x, scale_y)
                
            else:
                # 使用預設繪製
                self._draw_default_arc(draw, length, radius, img_width, img_height)
            
            return image
            
        except Exception as e:
            print(f"❌ 從 SVG 創建 type18 圖片失敗: {e}")
            return None
    
    def _draw_arc_rebar(self, draw, length, radius, img_width, img_height, scale_x, scale_y):
        """繪製圓弧鋼筋"""
        # 計算圓弧參數
        center_x = img_width // 2
        center_y = img_height // 2
        
        # 根據長度計算圓弧角度（假設圓弧長度對應角度）
        # 圓弧長度 = 半徑 × 角度（弧度）
        # 角度 = 圓弧長度 / 半徑
        arc_angle_rad = length / radius if radius > 0 else math.pi / 2
        arc_angle_deg = math.degrees(arc_angle_rad)
        
        # 限制角度範圍（0-180度）
        arc_angle_deg = min(arc_angle_deg, 180)
        
        # 計算圓弧的起始和結束角度
        start_angle = -arc_angle_deg / 2
        end_angle = arc_angle_deg / 2
        
        # 繪製圓弧
        line_width = 8
        bbox = [
            center_x - radius * scale_x,
            center_y - radius * scale_y,
            center_x + radius * scale_x,
            center_y + radius * scale_y
        ]
        
        # 使用 PIL 的 arc 方法繪製圓弧
        draw.arc(bbox, start_angle, end_angle, fill='black', width=line_width)
        
        # 添加標註
        self._add_arc_annotations(draw, length, radius, center_x, center_y, img_width, img_height)
    
    def _draw_default_arc(self, draw, length, radius, img_width, img_height):
        """繪製預設圓弧"""
        center_x = img_width // 2
        center_y = img_height // 2
        
        # 計算圓弧角度
        arc_angle_rad = length / radius if radius > 0 else math.pi / 2
        arc_angle_deg = math.degrees(arc_angle_rad)
        arc_angle_deg = min(arc_angle_deg, 180)
        
        # 繪製圓弧
        line_width = 8
        bbox = [
            center_x - radius,
            center_y - radius,
            center_x + radius,
            center_y + radius
        ]
        
        start_angle = -arc_angle_deg / 2
        end_angle = arc_angle_deg / 2
        
        draw.arc(bbox, start_angle, end_angle, fill='black', width=line_width)
        
        # 添加標註
        self._add_arc_annotations(draw, length, radius, center_x, center_y, img_width, img_height)
    
    def _add_arc_annotations(self, draw, length, radius, center_x, center_y, img_width, img_height):
        """添加圓弧標註"""
        font = self.get_font(32)
        small_font = self.get_font(24)
        
        # 在圓弧上方顯示長度
        length_text = str(int(length))
        text_x = center_x
        text_y = center_y - radius - 50
        self.draw_text_centered(draw, length_text, text_x, text_y, font)
        
        # 在圓弧下方顯示半徑
        radius_text = f"半徑={radius}"
        text_x = center_x
        text_y = center_y + radius + 30
        self.draw_text_centered(draw, radius_text, text_x, text_y, small_font)
