"""
Type12 折料圖形生成器
"""

from .base_generator import BaseImageGenerator

class Type12ImageGenerator(BaseImageGenerator):
    """Type12 折料圖形生成器"""
    
    def __init__(self):
        super().__init__()
        self.rebar_type = 'type12'
    
    def get_material_prefix(self):
        """獲取材料目錄前綴"""
        return "12-"
    
    def generate_image(self, segments, angles, rebar_number, available_materials):
        """生成 type12 鋼筋圖片"""
        try:
            # 尋找 type12 材料
            type12_material = self.find_material(available_materials)
            if not type12_material:
                print(f"❌ 找不到 type12 材料")
                return None
            
            # 構建 SVG 檔案路徑
            svg_path = self.get_svg_path(type12_material)
            if not svg_path.exists():
                print(f"❌ SVG 檔案不存在: {svg_path}")
                return None
            
            # 解析 SVG 並生成圖片
            return self._create_image_from_svg(svg_path, segments, angles, rebar_number)
            
        except Exception as e:
            print(f"❌ 生成 type12 鋼筋圖片失敗: {e}")
            return None
    
    def _create_image_from_svg(self, svg_path, segments, angles, rebar_number):
        """從 SVG 創建 type12 鋼筋圖片"""
        try:
            # 解析 SVG
            root = self.parse_svg(svg_path)
            if root is None:
                return None
            
            # 創建圖片
            img_width = 800
            img_height = 400
            image, draw = self.create_base_image(img_width, img_height)
            
            # 解析 SVG 中的 line 元素
            line_elements = root.findall(".//{http://www.w3.org/2000/svg}line")
            
            if len(line_elements) >= 2:
                # 計算縮放比例
                scale_x = img_width / 800
                scale_y = img_height / 600
                
                # 繪製鋼筋線條
                line_width = 8
                
                # 第一條線：水平線
                line1 = line_elements[0]
                x1 = int(float(line1.get('x1', 550)) * scale_x)
                y1 = int(float(line1.get('y1', 400)) * scale_y)
                x2 = int(float(line1.get('x2', 50)) * scale_x)
                y2 = int(float(line1.get('y2', 400)) * scale_y)
                draw.line([(x1, y1), (x2, y2)], fill='black', width=line_width)
                
                # 第二條線：斜線
                line2 = line_elements[1]
                x3 = int(float(line2.get('x1', 550)) * scale_x)
                y3 = int(float(line2.get('y1', 400)) * scale_y)
                x4 = int(float(line2.get('x2', 750)) * scale_x)
                y4 = int(float(line2.get('y2', 200)) * scale_y)
                draw.line([(x3, y3), (x4, y4)], fill='black', width=line_width)
                
                # 添加標註
                self._add_annotations(draw, segments, angles, x1, y1, x2, y2, x3, y3, x4, y4)
                
            else:
                # 使用預設繪製
                self._draw_default_shape(draw, segments, angles, img_width, img_height)
            
            return image
            
        except Exception as e:
            print(f"❌ 從 SVG 創建 type12 圖片失敗: {e}")
            return None
    
    def _add_annotations(self, draw, segments, angles, x1, y1, x2, y2, x3, y3, x4, y4):
        """添加標註"""
        font = self.get_font(32)
        small_font = self.get_font(24)
        
        # 在水平線上方顯示第一段長度
        if len(segments) >= 1:
            length1_text = str(int(segments[0]))
            text_x = (x1 + x2) // 2
            text_y = y1 - 50
            self.draw_text_centered(draw, length1_text, text_x, text_y, font)
        
        # 在斜線旁邊顯示第二段長度
        if len(segments) >= 2:
            length2_text = str(int(segments[1]))
            text_x = x4 + 10
            text_y = (y3 + y4) // 2 - 10
            draw.text((text_x, text_y), length2_text, fill='black', font=small_font)
        
        # 在斜線上方顯示角度
        if len(angles) >= 1:
            angle_text = f"{angles[0]}°"
            text_x = (x3 + x4) // 2
            text_y = min(y3, y4) - 30
            self.draw_text_centered(draw, angle_text, text_x, text_y, small_font)
    
    def _draw_default_shape(self, draw, segments, angles, img_width, img_height):
        """繪製預設形狀"""
        padding = 100
        line_start_x = padding
        line_end_x = img_width - padding
        line_y = img_height // 2
        
        # 繪製水平線段
        draw.line([(line_start_x, line_y), (line_end_x - 100, line_y)], fill='black', width=8)
        
        # 繪製斜線段
        hook_start_x = line_end_x - 100
        hook_end_x = line_end_x - 50
        hook_height = 100
        
        draw.line([(hook_start_x, line_y), (hook_end_x, line_y - hook_height)], fill='black', width=8)
        
        # 添加標註
        font = self.get_font(32)
        small_font = self.get_font(24)
        
        # 顯示長度
        if len(segments) >= 1:
            length_text = str(int(segments[0]))
            text_x = (line_start_x + line_end_x - 100) // 2
            text_y = line_y - 50
            self.draw_text_centered(draw, length_text, text_x, text_y, font)
        
        # 顯示角度
        if len(angles) >= 1:
            angle_text = f"{angles[0]}°"
            text_x = (hook_start_x + hook_end_x) // 2
            text_y = line_y - hook_height - 30
            self.draw_text_centered(draw, angle_text, text_x, text_y, small_font)
