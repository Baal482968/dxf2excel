"""
Type19 直段+弧段圖形生成器
"""

from .base_generator import BaseImageGenerator
import math

class Type19ImageGenerator(BaseImageGenerator):
    """Type19 直段+弧段圖形生成器"""
    
    def __init__(self):
        super().__init__()
        self.rebar_type = 'type19'
    
    def get_material_prefix(self):
        """獲取材料目錄前綴"""
        return "19-"
    
    def generate_image(self, straight_length, arc_length, radius, rebar_number, available_materials):
        """生成 type19 鋼筋圖片"""
        try:
            # 尋找 type19 材料
            type19_material = self.find_material(available_materials)
            if not type19_material:
                print(f"❌ 找不到 type19 材料")
                return None
            
            # 構建 SVG 檔案路徑
            svg_path = self.get_svg_path(type19_material)
            if not svg_path.exists():
                print(f"❌ SVG 檔案不存在: {svg_path}")
                return None
            
            # 解析 SVG 並生成圖片
            return self._create_image_from_svg(svg_path, straight_length, arc_length, radius, rebar_number)
            
        except Exception as e:
            print(f"❌ 生成 type19 鋼筋圖片失敗: {e}")
            return None
    
    def _create_image_from_svg(self, svg_path, straight_length, arc_length, radius, rebar_number):
        """從 SVG 創建 type19 鋼筋圖片"""
        try:
            # 解析 SVG
            root = self.parse_svg(svg_path)
            if root is None:
                return None
            
            # 創建圖片
            img_width = 800
            img_height = 400
            image, draw = self.create_base_image(img_width, img_height)
            
            # 解析 SVG 中的元素
            line_element = root.find(".//{http://www.w3.org/2000/svg}line")
            path_element = root.find(".//{http://www.w3.org/2000/svg}path")
            
            if line_element is not None and path_element is not None:
                # 計算縮放比例
                scale_x = img_width / 800
                scale_y = img_height / 600
                
                # 繪製直段+弧段鋼筋
                self._draw_straight_arc_rebar(draw, straight_length, arc_length, radius, img_width, img_height, scale_x, scale_y)
                
            else:
                # 使用預設繪製
                self._draw_default_straight_arc(draw, straight_length, arc_length, radius, img_width, img_height)
            
            return image
            
        except Exception as e:
            print(f"❌ 從 SVG 創建 type19 圖片失敗: {e}")
            return None
    
    def _draw_straight_arc_rebar(self, draw, straight_length, arc_length, radius, img_width, img_height, scale_x, scale_y):
        """繪製直段+弧段鋼筋"""
        # 計算位置
        center_x = img_width // 2
        center_y = img_height // 2
        
        # 繪製直段（水平線）
        line_width = 8
        straight_start_x = int(50 * scale_x)
        straight_end_x = int(550 * scale_x)
        straight_y = int(375 * scale_y)
        
        draw.line([(straight_start_x, straight_y), (straight_end_x, straight_y)], fill='black', width=line_width)
        
        # 繪製弧段
        # 計算弧段的角度
        arc_angle_rad = arc_length / radius if radius > 0 else math.pi / 2
        arc_angle_deg = math.degrees(arc_angle_rad)
        
        # 弧段起始點（直段終點）
        arc_start_x = straight_end_x
        arc_start_y = straight_y
        
        # 計算弧段中心點（在弧段起始點上方）
        arc_center_x = arc_start_x
        arc_center_y = arc_start_y - int(radius * scale_y)
        
        # 繪製弧段
        start_angle = 0
        end_angle = arc_angle_deg
        
        # 使用橢圓弧繪製
        bbox = [
            arc_center_x - int(radius * scale_x),
            arc_center_y - int(radius * scale_y),
            arc_center_x + int(radius * scale_x),
            arc_center_y + int(radius * scale_y)
        ]
        
        # 繪製弧線
        draw.arc(bbox, start_angle, end_angle, fill='black', width=line_width)
        
        # 添加標註
        self._add_straight_arc_annotations(draw, straight_length, arc_length, radius, center_x, center_y, img_width, img_height)
    
    def _draw_default_straight_arc(self, draw, straight_length, arc_length, radius, img_width, img_height):
        """繪製預設的直段+弧段鋼筋"""
        # 計算位置
        center_x = img_width // 2
        center_y = img_height // 2
        
        # 繪製直段
        line_width = 8
        straight_start_x = 50
        straight_end_x = 550
        straight_y = 375
        
        draw.line([(straight_start_x, straight_y), (straight_end_x, straight_y)], fill='black', width=line_width)
        
        # 繪製弧段
        arc_angle_rad = arc_length / radius if radius > 0 else math.pi / 2
        arc_angle_deg = math.degrees(arc_angle_rad)
        
        arc_center_x = straight_end_x
        arc_center_y = straight_y - radius
        
        bbox = [
            arc_center_x - radius,
            arc_center_y - radius,
            arc_center_x + radius,
            arc_center_y + radius
        ]
        
        draw.arc(bbox, 0, arc_angle_deg, fill='black', width=line_width)
        
        # 添加標註
        self._add_straight_arc_annotations(draw, straight_length, arc_length, radius, center_x, center_y, img_width, img_height)
    
    def _add_straight_arc_annotations(self, draw, straight_length, arc_length, radius, center_x, center_y, img_width, img_height):
        """添加直段+弧段標註"""
        font = self.get_font(32)
        small_font = self.get_font(24)
        
        # 在直段下方顯示長度
        straight_text = str(int(straight_length))
        text_x = center_x
        text_y = 375 + 20  # 在直段下方
        self.draw_text_centered(draw, straight_text, text_x, text_y, font)
        
        # 在弧段右方顯示長度
        arc_text = str(int(arc_length))
        text_x = 550 + 20  # 在弧段右方
        text_y = center_y - 20
        draw.text((text_x, text_y), arc_text, fill='black', font=font)
        
        # 在弧段上方顯示半徑
        radius_text = f"半徑 {radius}"
        text_x = center_x
        text_y = center_y - 80  # 在弧段上方
        self.draw_text_centered(draw, radius_text, text_x, text_y, small_font)
