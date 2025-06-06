"""
圖形繪製相關功能模組 - 增強版
整合專業鋼筋圖示功能，符合範例1的需求
"""

import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="matplotlib")

# 設定中文字體支援
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'Heiti TC', 'Microsoft JhengHei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

import matplotlib.patches as patches
from matplotlib.patches import Arc
import numpy as np
from PIL import Image
import io
import base64
import os

class GraphicsManager:
    """圖形繪製管理器 - 增強版"""
    
    def __init__(self):
        """初始化圖形管理器"""
        # 標準鋼筋直徑對照表 (mm) - 符合CNS 560
        self.rebar_diameters = {
            "#2": 6.4, "#3": 9.5, "#4": 12.7, "#5": 15.9, "#6": 19.1,
            "#7": 22.2, "#8": 25.4, "#9": 28.7, "#10": 32.3, "#11": 35.8,
            "#12": 43.0, "#13": 50.0, "#14": 57.0
        }
        
        # 標準彎曲半徑倍數 - 依據建築技術規則
        self.bend_radius_multiplier = {
            "#2": 3, "#3": 3, "#4": 4, "#5": 5, "#6": 6, "#7": 7,
            "#8": 8, "#9": 9, "#10": 10, "#11": 11, "#12": 12,
            "#13": 13, "#14": 14
        }
        
        # 專業模式圖形參數
        self.professional_settings = {
            'line_width': 4,
            'font_size': 12,
            'dimension_offset': 25,
            'margin': 30,
            'colors': {
                'rebar': '#2C3E50',      # 鋼筋顏色（深藍灰）
                'dimension': '#E74C3C',   # 尺寸線顏色（紅色）
                'text': '#2C3E50',       # 文字顏色
                'radius': '#27AE60'      # 半徑標註顏色（綠色）
            }
        }
        
        # 基本模式圖形參數
        self.basic_settings = {
            'line_width': 2,
            'font_size': 10,
            'colors': {
                'rebar': 'blue',
                'text': 'black'
            }
        }
    
    @staticmethod
    def check_dependencies():
        """檢查圖形繪製所需的套件是否已安裝"""
        try:
            import matplotlib
            import numpy
            from PIL import Image
            return True, []
        except ImportError as e:
            missing_packages = []
            
            try:
                import matplotlib
            except ImportError:
                missing_packages.append("matplotlib")
            
            try:
                import numpy
            except ImportError:
                missing_packages.append("numpy")
            
            try:
                from PIL import Image
            except ImportError:
                missing_packages.append("pillow")
            
            return False, missing_packages
    
    def get_bend_radius(self, rebar_number):
        """計算彎曲半徑（符合建築技術規則）"""
        diameter = self.rebar_diameters.get(rebar_number, 12.7)
        multiplier = self.bend_radius_multiplier.get(rebar_number, 4)
        return diameter * multiplier
    
    def get_material_grade(self, rebar_number):
        """根據鋼筋編號判定材料等級"""
        if rebar_number in ["#2", "#3", "#4", "#5", "#6"]:
            return "SD280"
        elif rebar_number in ["#7", "#8", "#9", "#10"]:
            return "SD420"
        else:
            return "SD490"
    
    def draw_dimension_line(self, ax, start_point, end_point, offset, text, 
                           horizontal=True, settings=None):
        """繪製專業尺寸標註線"""
        if settings is None:
            settings = self.professional_settings
            
        x1, y1 = start_point
        x2, y2 = end_point
        color = settings['colors']['dimension']
        font_size = settings['font_size']
        
        if horizontal:
            # 水平尺寸線
            dim_y = y1 + offset
            # 主尺寸線
            ax.plot([x1, x2], [dim_y, dim_y], color=color, linewidth=1.5)
            # 引出線
            ax.plot([x1, x1], [y1, dim_y], color=color, linewidth=1)
            ax.plot([x2, x2], [y2, dim_y], color=color, linewidth=1)
            # 箭頭
            arrow_size = 3
            ax.plot([x1, x1+arrow_size], [dim_y, dim_y-arrow_size/2], color=color, linewidth=1.5)
            ax.plot([x1, x1+arrow_size], [dim_y, dim_y+arrow_size/2], color=color, linewidth=1.5)
            ax.plot([x2, x2-arrow_size], [dim_y, dim_y-arrow_size/2], color=color, linewidth=1.5)
            ax.plot([x2, x2-arrow_size], [dim_y, dim_y+arrow_size/2], color=color, linewidth=1.5)
            # 尺寸文字
            mid_x = (x1 + x2) / 2
            ax.text(mid_x, dim_y + 8, text, ha='center', va='bottom',
                   fontsize=font_size, fontweight='bold', color=settings['colors']['text'])
        else:
            # 垂直尺寸線
            dim_x = x1 + offset
            ax.plot([dim_x, dim_x], [y1, y2], color=color, linewidth=1.5)
            ax.plot([x1, dim_x], [y1, y1], color=color, linewidth=1)
            ax.plot([x2, dim_x], [y2, y2], color=color, linewidth=1)
            # 箭頭
            arrow_size = 3
            ax.plot([dim_x, dim_x-arrow_size/2], [y1, y1+arrow_size], color=color, linewidth=1.5)
            ax.plot([dim_x, dim_x+arrow_size/2], [y1, y1+arrow_size], color=color, linewidth=1.5)
            ax.plot([dim_x, dim_x-arrow_size/2], [y2, y2-arrow_size], color=color, linewidth=1.5)
            ax.plot([dim_x, dim_x+arrow_size/2], [y2, y2-arrow_size], color=color, linewidth=1.5)
            # 尺寸文字
            mid_y = (y1 + y2) / 2
            ax.text(dim_x + 12, mid_y, text, ha='left', va='center',
                   fontsize=font_size, fontweight='bold', color=settings['colors']['text'], rotation=90)
    
    def draw_straight_rebar(self, length, rebar_number, professional=True):
        """繪製直鋼筋圖示"""
        if professional:
            return self._draw_professional_straight_rebar(length, rebar_number)
        else:
            return self._draw_basic_straight_rebar(length, rebar_number)
    
    def _draw_professional_straight_rebar(self, length, rebar_number):
        """繪製專業直鋼筋圖示"""
        fig, ax = plt.subplots(figsize=(10, 4))
        ax.set_aspect('equal')
        
        settings = self.professional_settings
        scale = 3  # 1cm = 3 units
        length_scaled = length * scale
        
        # 起始點
        start_x, start_y = 50, 50
        end_x = start_x + length_scaled
        
        # 繪製鋼筋主線
        ax.plot([start_x, end_x], [start_y, start_y], 
                color=settings['colors']['rebar'], 
                linewidth=settings['line_width'], 
                solid_capstyle='round')
        
        # 繪製尺寸標註
        self.draw_dimension_line(ax, (start_x, start_y), (end_x, start_y), 
                                settings['dimension_offset'], f'{int(length)}', 
                                horizontal=True, settings=settings)
        
        # 鋼筋編號標註
        ax.text(start_x - 25, start_y, rebar_number, 
                ha='center', va='center', 
                fontsize=settings['font_size'] + 4, 
                fontweight='bold', 
                color=settings['colors']['text'],
                bbox=dict(boxstyle="round,pad=0.3", facecolor='white', edgecolor='gray'))
        
        # 技術資訊標註
        diameter = self.rebar_diameters.get(rebar_number, 12.7)
        material_grade = self.get_material_grade(rebar_number)
        ax.text(end_x + 30, start_y + 20, 
                f'D{diameter}mm\n{material_grade}', 
                ha='left', va='center', 
                fontsize=settings['font_size'] - 1,
                style='italic', 
                color=settings['colors']['text'])
        
        # 設定圖形範圍
        ax.set_xlim(start_x - 60, end_x + 80)
        ax.set_ylim(start_y - 40, start_y + settings['dimension_offset'] + 40)
        ax.axis('off')
        
        # 添加標題
        fig.suptitle(f'直鋼筋 {rebar_number} - 長度 {int(length)}cm', 
                    fontsize=settings['font_size'] + 2, fontweight='bold')
        
        plt.tight_layout()
        return self._figure_to_base64(fig)
    
    def _draw_basic_straight_rebar(self, length, rebar_number):
        """繪製基本直鋼筋圖示（保留原有功能）"""
        fig, ax = plt.subplots(figsize=(6, 1))
        ax.plot([0, length], [0, 0], 
                color=self.basic_settings['colors']['rebar'], 
                linewidth=self.basic_settings['line_width'])
        ax.set_xlim(-1, length + 1)
        ax.set_ylim(-1, 1)
        ax.axis('off')
        
        # 加入鋼筋編號標記
        ax.text(length/2, 0.3, rebar_number, ha='center', va='center')
        
        return self._figure_to_base64(fig)

    def draw_l_shaped_rebar(self, length1, length2, rebar_number, professional=True):
        """繪製 L 型鋼筋圖示"""
        if professional:
            return self._draw_professional_l_shaped_rebar(length1, length2, rebar_number)
        else:
            return self._draw_basic_l_shaped_rebar(length1, length2, rebar_number)
    
    def _draw_professional_l_shaped_rebar(self, length1, length2, rebar_number):
        """繪製專業L型鋼筋圖示"""
        fig, ax = plt.subplots(figsize=(10, 8))
        ax.set_aspect('equal')
        
        settings = self.professional_settings
        scale = 3
        length1_scaled = length1 * scale
        length2_scaled = length2 * scale
        bend_radius = self.get_bend_radius(rebar_number) / 10 * scale
        
        # 起始點
        start_x, start_y = 60, 120
        
        # 第一段 (水平)
        h1_end_x = start_x + length1_scaled
        ax.plot([start_x, h1_end_x - bend_radius], [start_y, start_y], 
                color=settings['colors']['rebar'], 
                linewidth=settings['line_width'], 
                solid_capstyle='round')
        
        # 彎曲部分 (90度弧)
        corner_center_x = h1_end_x - bend_radius
        corner_center_y = start_y + bend_radius
        
        arc = Arc((corner_center_x, corner_center_y), 
                  2 * bend_radius, 2 * bend_radius,
                  angle=0, theta1=270, theta2=360,
                  linewidth=settings['line_width'], 
                  color=settings['colors']['rebar'])
        ax.add_patch(arc)
        
        # 第二段 (垂直)
        v_start_x = h1_end_x
        v_start_y = start_y + bend_radius
        v_end_y = v_start_y + length2_scaled
        ax.plot([v_start_x, v_start_x], [v_start_y, v_end_y], 
                color=settings['colors']['rebar'], 
                linewidth=settings['line_width'], 
                solid_capstyle='round')
        
        # 尺寸標註
        self.draw_dimension_line(ax, (start_x, start_y), (h1_end_x, start_y), 
                                -settings['dimension_offset'], f'{int(length1)}', 
                                horizontal=True, settings=settings)
        
        self.draw_dimension_line(ax, (v_start_x, v_start_y), (v_start_x, v_end_y), 
                                settings['dimension_offset'], f'{int(length2)}', 
                                horizontal=False, settings=settings)
        
        # 鋼筋編號標註
        ax.text(start_x - 35, start_y, rebar_number, 
                ha='center', va='center', 
                fontsize=settings['font_size'] + 4, 
                fontweight='bold', 
                color=settings['colors']['text'],
                bbox=dict(boxstyle="round,pad=0.3", facecolor='white', edgecolor='gray'))
        
        # 彎曲半徑標註
        radius_text = f'R{int(bend_radius*10/scale)}'
        ax.text(corner_center_x + bend_radius/2, corner_center_y - bend_radius/2, 
                radius_text, ha='center', va='center', 
                fontsize=settings['font_size'] - 1, 
                color=settings['colors']['radius'],
                bbox=dict(boxstyle="round,pad=0.2", facecolor='lightyellow', alpha=0.8))
        
        # 技術資訊
        diameter = self.rebar_diameters.get(rebar_number, 12.7)
        material_grade = self.get_material_grade(rebar_number)
        total_length = length1 + length2
        ax.text(v_start_x + 50, v_end_y, 
                f'D{diameter}mm\n{material_grade}\n總長: {int(total_length)}cm', 
                ha='left', va='top', 
                fontsize=settings['font_size'] - 1,
                color=settings['colors']['text'], 
                linespacing=1.5)
        
        # 設定圖形範圍
        ax.set_xlim(start_x - 70, max(h1_end_x, v_start_x + settings['dimension_offset']) + 80)
        ax.set_ylim(start_y - settings['dimension_offset'] - 30, v_end_y + 30)
        ax.axis('off')
        
        # 添加標題
        fig.suptitle(f'L型鋼筋 {rebar_number} - {int(length1)}cm + {int(length2)}cm', 
                    fontsize=settings['font_size'] + 2, fontweight='bold')
        
        plt.tight_layout()
        return self._figure_to_base64(fig)
    
    def _draw_basic_l_shaped_rebar(self, length1, length2, rebar_number):
        """繪製基本L型鋼筋圖示（保留原有功能）"""
        fig, ax = plt.subplots(figsize=(6, 3))
        
        # 繪製 L 型鋼筋
        ax.plot([0, length1], [0, 0], 
                color=self.basic_settings['colors']['rebar'], 
                linewidth=self.basic_settings['line_width'])
        ax.plot([length1, length1], [0, -length2], 
                color=self.basic_settings['colors']['rebar'], 
                linewidth=self.basic_settings['line_width'])
        
        ax.set_xlim(-1, length1 + 1)
        ax.set_ylim(-length2 - 1, 1)
        ax.axis('off')
        
        # 加入鋼筋編號標記
        ax.text(length1/2, 0.3, rebar_number, ha='center', va='center')
        
        return self._figure_to_base64(fig)

    def draw_u_shaped_rebar(self, length1, length2, length3, rebar_number, professional=True):
        """繪製 U 型鋼筋圖示"""
        if professional:
            return self._draw_professional_u_shaped_rebar(length1, length2, length3, rebar_number)
        else:
            return self._draw_basic_u_shaped_rebar(length1, length2, length3, rebar_number)
    
    def _draw_professional_u_shaped_rebar(self, length1, length2, length3, rebar_number):
        """繪製專業U型鋼筋圖示"""
        fig, ax = plt.subplots(figsize=(12, 8))
        ax.set_aspect('equal')
        
        settings = self.professional_settings
        scale = 2.5
        length1_scaled = length1 * scale
        length2_scaled = length2 * scale
        length3_scaled = length3 * scale
        bend_radius = self.get_bend_radius(rebar_number) / 10 * scale
        
        # 起始點
        start_x, start_y = 80, 150
        
        # 第一段 (垂直向下)
        v1_end_y = start_y - length1_scaled
        ax.plot([start_x, start_x], [start_y, v1_end_y + bend_radius], 
                color=settings['colors']['rebar'], 
                linewidth=settings['line_width'], 
                solid_capstyle='round')
        
        # 第一個彎曲 (左下角)
        corner1_center_x = start_x + bend_radius
        corner1_center_y = v1_end_y + bend_radius
        arc1 = Arc((corner1_center_x, corner1_center_y), 
                   2 * bend_radius, 2 * bend_radius,
                   angle=0, theta1=180, theta2=270,
                   linewidth=settings['line_width'], 
                   color=settings['colors']['rebar'])
        ax.add_patch(arc1)
        
        # 第二段 (水平)
        h_start_x = start_x + bend_radius
        h_end_x = h_start_x + length2_scaled
        ax.plot([h_start_x, h_end_x - bend_radius], [v1_end_y, v1_end_y], 
                color=settings['colors']['rebar'], 
                linewidth=settings['line_width'], 
                solid_capstyle='round')
        
        # 第二個彎曲 (右下角)
        corner2_center_x = h_end_x - bend_radius
        corner2_center_y = v1_end_y + bend_radius
        arc2 = Arc((corner2_center_x, corner2_center_y), 
                   2 * bend_radius, 2 * bend_radius,
                   angle=0, theta1=270, theta2=360,
                   linewidth=settings['line_width'], 
                   color=settings['colors']['rebar'])
        ax.add_patch(arc2)
        
        # 第三段 (垂直向上)
        v2_start_y = v1_end_y + bend_radius
        v2_end_y = v2_start_y + length3_scaled
        ax.plot([h_end_x, h_end_x], [v2_start_y, v2_end_y], 
                color=settings['colors']['rebar'], 
                linewidth=settings['line_width'], 
                solid_capstyle='round')
        
        # 尺寸標註
        self.draw_dimension_line(ax, (start_x, start_y), (start_x, v1_end_y), 
                                -settings['dimension_offset'], f'{int(length1)}', 
                                horizontal=False, settings=settings)
        
        self.draw_dimension_line(ax, (start_x, v1_end_y), (h_end_x, v1_end_y), 
                                -settings['dimension_offset'], f'{int(length2)}', 
                                horizontal=True, settings=settings)
        
        self.draw_dimension_line(ax, (h_end_x, v1_end_y), (h_end_x, v2_end_y), 
                                settings['dimension_offset'], f'{int(length3)}', 
                                horizontal=False, settings=settings)
        
        # 鋼筋編號和技術資訊
        ax.text(start_x - 45, start_y, rebar_number, 
                ha='center', va='center', 
                fontsize=settings['font_size'] + 4, 
                fontweight='bold', 
                color=settings['colors']['text'],
                bbox=dict(boxstyle="round,pad=0.3", facecolor='white', edgecolor='gray'))
        
        # 彎曲半徑標註
        radius_text = f'R{int(bend_radius*10/scale)}'
        ax.text(corner1_center_x - bend_radius/2, corner1_center_y - bend_radius/2, 
                radius_text, ha='center', va='center', 
                fontsize=settings['font_size'] - 1, 
                color=settings['colors']['radius'],
                bbox=dict(boxstyle="round,pad=0.2", facecolor='lightyellow', alpha=0.8))
        
        # 技術資訊
        diameter = self.rebar_diameters.get(rebar_number, 12.7)
        material_grade = self.get_material_grade(rebar_number)
        total_length = length1 + length2 + length3
        ax.text(h_end_x + 50, v2_end_y, 
                f'D{diameter}mm\n{material_grade}\n總長: {int(total_length)}cm\nU型鋼筋', 
                ha='left', va='top', 
                fontsize=settings['font_size'] - 1,
                color=settings['colors']['text'], 
                linespacing=1.5)
        
        # 設定圖形範圍
        ax.set_xlim(start_x - settings['dimension_offset'] - 60, 
                   h_end_x + settings['dimension_offset'] + 80)
        ax.set_ylim(v1_end_y - settings['dimension_offset'] - 30, start_y + 30)
        ax.axis('off')
        
        # 添加標題
        fig.suptitle(f'U型鋼筋 {rebar_number} - {int(length1)} + {int(length2)} + {int(length3)}cm', 
                    fontsize=settings['font_size'] + 2, fontweight='bold')
        
        plt.tight_layout()
        return self._figure_to_base64(fig)
    
    def _draw_basic_u_shaped_rebar(self, length1, length2, length3, rebar_number):
        """繪製基本U型鋼筋圖示（保留原有功能）"""
        fig, ax = plt.subplots(figsize=(6, 3))
        
        # 繪製 U 型鋼筋
        ax.plot([0, length1], [0, 0], 
                color=self.basic_settings['colors']['rebar'], 
                linewidth=self.basic_settings['line_width'])
        ax.plot([length1, length1], [0, -length2], 
                color=self.basic_settings['colors']['rebar'], 
                linewidth=self.basic_settings['line_width'])
        ax.plot([length1, 0], [-length2, -length2], 
                color=self.basic_settings['colors']['rebar'], 
                linewidth=self.basic_settings['line_width'])
        ax.plot([0, 0], [-length2, 0], 
                color=self.basic_settings['colors']['rebar'], 
                linewidth=self.basic_settings['line_width'])
        
        ax.set_xlim(-1, length1 + 1)
        ax.set_ylim(-length2 - 1, 1)
        ax.axis('off')
        
        # 加入鋼筋編號標記
        ax.text(length1/2, 0.3, rebar_number, ha='center', va='center')
        
        return self._figure_to_base64(fig)

    def draw_complex_rebar(self, segments, rebar_number, professional=True):
        """繪製複雜多段鋼筋圖示"""
        if professional:
            return self._draw_professional_complex_rebar(segments, rebar_number)
        else:
            return self._draw_basic_complex_rebar(segments, rebar_number)
    
    def _draw_professional_complex_rebar(self, segments, rebar_number):
        """繪製專業複雜鋼筋圖示"""
        fig, ax = plt.subplots(figsize=(14, 10))
        ax.set_aspect('equal')
        
        settings = self.professional_settings
        scale = 2
        bend_radius = self.get_bend_radius(rebar_number) / 10 * scale
        
        # 起始點
        current_x, current_y = 100, 150
        start_x, start_y = current_x, current_y
        
        # 方向序列：右、下、左、上，循環
        directions = [(1, 0), (0, -1), (-1, 0), (0, 1)]
        current_direction = 0
        
        points = [(current_x, current_y)]
        segment_midpoints = []
        
        # 繪製各段
        for i, length in enumerate(segments):
            length_scaled = length * scale
            dx, dy = directions[current_direction % 4]
            
            if i > 0:  # 非第一段需要彎曲
                # 計算彎曲圓心
                prev_dx, prev_dy = directions[(current_direction - 1) % 4]
                center_x = current_x + bend_radius * prev_dx + bend_radius * dx
                center_y = current_y + bend_radius * prev_dy + bend_radius * dy
                
                # 計算弧的角度
                start_angle = np.degrees(np.arctan2(-prev_dy, -prev_dx))
                end_angle = np.degrees(np.arctan2(-dy, -dx))
                
                # 調整角度範圍
                angle_diff = (end_angle - start_angle) % 360
                if angle_diff > 180:
                    start_angle, end_angle = end_angle, start_angle + 360
                
                # 繪製弧線
                arc = Arc((center_x, center_y), 
                         2 * bend_radius, 2 * bend_radius,
                         angle=0, theta1=start_angle, theta2=end_angle,
                         linewidth=settings['line_width'], 
                         color=settings['colors']['rebar'])
                ax.add_patch(arc)
                
                # 彎曲半徑標註（僅標註前兩個彎曲）
                if i <= 2:
                    radius_text = f'R{int(bend_radius*10/scale)}'
                    offset_x = bend_radius/3 * (prev_dx + dx)
                    offset_y = bend_radius/3 * (prev_dy + dy)
                    ax.text(center_x + offset_x, center_y + offset_y, 
                            radius_text, ha='center', va='center', 
                            fontsize=settings['font_size'] - 2, 
                            color=settings['colors']['radius'],
                            bbox=dict(boxstyle="round,pad=0.1", facecolor='lightyellow', alpha=0.7))
                
                # 更新當前位置
                current_x = center_x - bend_radius * dx
                current_y = center_y - bend_radius * dy
            
            # 計算段終點
            segment_end_x = current_x + length_scaled * dx
            segment_end_y = current_y + length_scaled * dy
            
            # 如果不是最後一段，要為下一個彎曲預留空間
            if i < len(segments) - 1:
                draw_end_x = segment_end_x - bend_radius * dx
                draw_end_y = segment_end_y - bend_radius * dy
            else:
                draw_end_x = segment_end_x
                draw_end_y = segment_end_y
            
            # 繪製直線段
            ax.plot([current_x, draw_end_x], [current_y, draw_end_y], 
                    color=settings['colors']['rebar'], 
                    linewidth=settings['line_width'], 
                    solid_capstyle='round')
            
            # 記錄段中點用於標註
            mid_x = (current_x + draw_end_x) / 2
            mid_y = (current_y + draw_end_y) / 2
            segment_midpoints.append((mid_x, mid_y, length, dx, dy, i))
            
            # 更新當前位置
            current_x, current_y = segment_end_x, segment_end_y
            points.append((current_x, current_y))
            current_direction += 1
        
        # 添加尺寸標註
        for mid_x, mid_y, length, dx, dy, i in segment_midpoints:
            # 計算標註位置偏移
            if dx != 0:  # 水平線段
                offset_y = settings['dimension_offset'] if i % 2 == 0 else -settings['dimension_offset']
                ax.text(mid_x, mid_y + offset_y, f'{int(length)}', 
                        ha='center', va='bottom' if offset_y > 0 else 'top', 
                        fontsize=settings['font_size'], fontweight='bold', 
                        color=settings['colors']['text'],
                        bbox=dict(boxstyle="round,pad=0.2", facecolor='white', alpha=0.8))
            else:  # 垂直線段
                offset_x = settings['dimension_offset'] if i % 2 == 0 else -settings['dimension_offset']
                ax.text(mid_x + offset_x, mid_y, f'{int(length)}', 
                        ha='left' if offset_x > 0 else 'right', va='center', 
                        fontsize=settings['font_size'], fontweight='bold', 
                        color=settings['colors']['text'], rotation=90,
                        bbox=dict(boxstyle="round,pad=0.2", facecolor='white', alpha=0.8))
        
        # 鋼筋編號標註
        ax.text(start_x - 60, start_y, rebar_number, 
                ha='center', va='center', 
                fontsize=settings['font_size'] + 4, 
                fontweight='bold', 
                color=settings['colors']['text'],
                bbox=dict(boxstyle="round,pad=0.4", facecolor='white', edgecolor='gray'))
        
        # 技術資訊
        diameter = self.rebar_diameters.get(rebar_number, 12.7)
        material_grade = self.get_material_grade(rebar_number)
        total_length = sum(segments)
        segment_text = ' + '.join([str(int(s)) for s in segments])
        
        # 找到圖形右上角位置放置資訊
        all_x = [p[0] for p in points]
        all_y = [p[1] for p in points]
        info_x = max(all_x) + 30
        info_y = max(all_y)
        
        ax.text(info_x, info_y, 
                f'鋼筋規格: D{diameter}mm\n'
                f'材料等級: {material_grade}\n'
                f'總長度: {int(total_length)}cm\n'
                f'分段: {segment_text}cm\n'
                f'段數: {len(segments)}段\n'
                f'形狀: 複合彎曲', 
                ha='left', va='top', 
                fontsize=settings['font_size'] - 1,
                color=settings['colors']['text'], 
                linespacing=1.5,
                bbox=dict(boxstyle="round,pad=0.5", facecolor='lightblue', alpha=0.3))
        
        # 自動調整圖形範圍
        margin = 80
        ax.set_xlim(min(all_x) - margin, max(all_x) + margin + 150)
        ax.set_ylim(min(all_y) - margin, max(all_y) + margin)
        ax.axis('off')
        
        # 添加標題
        fig.suptitle(f'複合鋼筋 {rebar_number} - {len(segments)}段彎曲', 
                    fontsize=settings['font_size'] + 2, fontweight='bold')
        
        plt.tight_layout()
        return self._figure_to_base64(fig)
    
    def _draw_basic_complex_rebar(self, segments, rebar_number):
        """繪製基本複雜鋼筋圖示（保留原有功能）"""
        fig, ax = plt.subplots(figsize=(8, 4))
        
        # 計算總長度和高度
        total_length = sum(segments)
        max_height = max(segments)
        
        # 繪製多段鋼筋
        x, y = 0, 0
        for i, length in enumerate(segments):
            if i % 2 == 0:
                ax.plot([x, x + length], [y, y], 
                        color=self.basic_settings['colors']['rebar'], 
                        linewidth=self.basic_settings['line_width'])
                x += length
            else:
                ax.plot([x, x], [y, y - length], 
                        color=self.basic_settings['colors']['rebar'], 
                        linewidth=self.basic_settings['line_width'])
                y -= length
        
        ax.set_xlim(-1, total_length + 1)
        ax.set_ylim(-max_height - 1, 1)
        ax.axis('off')
        
        # 加入鋼筋編號標記
        ax.text(total_length/2, 0.3, rebar_number, ha='center', va='center')
        
        return self._figure_to_base64(fig)

    def draw_ascii_rebar(self, segments):
        """生成 ASCII 格式的鋼筋圖示（保留原有功能）"""
        if not segments:
            return ""
            
        # 根據段數生成不同的ASCII圖示
        if len(segments) == 1:
            # 直鋼筋
            length = int(segments[0])
            return f"{'─' * min(length//10, 20)}"
        
        elif len(segments) == 2:
            # L型鋼筋
            return """┌─────────
│
│
│
└"""
        
        elif len(segments) == 3:
            # U型鋼筋
            return """│     │
│     │
└─────┘"""
        
        else:
            # 複雜鋼筋
            return """┌─┬─┐
│ │ │
└─┴─┘"""

    def create_detailed_description(self, segments, rebar_number):
        """創建詳細的鋼筋描述文字"""
        # 基本資訊計算
        total_length = sum(segments)
        diameter = self.rebar_diameters.get(rebar_number, 12.7)
        bend_radius = self.get_bend_radius(rebar_number)
        material_grade = self.get_material_grade(rebar_number)
        
        # 確定鋼筋類型
        if len(segments) == 1:
            shape_type = "直鋼筋"
            bend_info = "無彎曲"
        elif len(segments) == 2:
            shape_type = "L型鋼筋"
            bend_info = "90°彎曲 × 1"
        elif len(segments) == 3:
            shape_type = "U型鋼筋"
            bend_info = "90°彎曲 × 2"
        else:
            shape_type = f"{len(segments)}段複合鋼筋"
            bend_info = f"90°彎曲 × {len(segments)-1}"
        
        # 分段資訊
        segment_info = ""
        total_with_letters = ""
        for i, length in enumerate(segments):
            letter = chr(65 + i)  # A, B, C, ...
            segment_info += f"    {letter}段: {int(length):>4} cm\n"
            if i == 0:
                total_with_letters = f"{int(length)}"
            else:
                total_with_letters += f" + {int(length)}"
        
        # 組合完整描述
        description = f"""╔══════════════════════════════════════════════════════════╗
║                    專業鋼筋規格圖                        ║
╚══════════════════════════════════════════════════════════╝

🔧 基本資訊:
  鋼筋編號: {rebar_number}
  鋼筋類型: {shape_type}
  材料等級: {material_grade}

📏 尺寸規格:
  總長度: {int(total_length):>4} cm
  計算式: {total_with_letters} = {int(total_length)} cm
{segment_info}
⚙️ 技術參數:
  鋼筋直徑: D{diameter} mm
  彎曲半徑: R{int(bend_radius)} mm
  彎曲資訊: {bend_info}
  最小保護層: {int(diameter + 10)} mm

📋 設計依據:
  • CNS 560 鋼筋混凝土用鋼筋
  • 建築技術規則施工編
  • 結構混凝土設計規範
  • 最小彎曲半徑 = {bend_radius//diameter:.0f} × D

⚠️  施工注意事項:
  • 彎曲時須使用標準彎曲機具
  • 彎曲半徑不得小於規範要求
  • 彎曲角度誤差不得超過±2°
  • 鋼筋表面不得有裂縫或損傷"""
        
        return description

    def generate_rebar_diagram(self, segments, rebar_number, mode="professional"):
        """
        主要入口函數，根據段數和模式生成對應的鋼筋圖示
        
        Args:
            segments: 分段長度列表 (cm)
            rebar_number: 鋼筋編號
            mode: 繪圖模式 ("professional", "basic", "ascii")
        
        Returns:
            str: base64編碼的圖片或ASCII文字
        """
        if not segments:
            return self.draw_ascii_rebar([])
        
        # 過濾有效分段
        valid_segments = [s for s in segments if s and s > 0]
        if not valid_segments:
            return self.draw_ascii_rebar([])
        
        try:
            if mode == "ascii":
                return self.draw_ascii_rebar(valid_segments)
            
            professional = (mode == "professional")
            
            if len(valid_segments) == 1:
                return self.draw_straight_rebar(valid_segments[0], rebar_number, professional)
            elif len(valid_segments) == 2:
                return self.draw_l_shaped_rebar(valid_segments[0], valid_segments[1], rebar_number, professional)
            elif len(valid_segments) == 3:
                return self.draw_u_shaped_rebar(valid_segments[0], valid_segments[1], valid_segments[2], rebar_number, professional)
            else:
                return self.draw_complex_rebar(valid_segments, rebar_number, professional)
                
        except Exception as e:
            # 如果專業模式失敗，嘗試基本模式
            if mode == "professional":
                try:
                    return self.generate_rebar_diagram(segments, rebar_number, "basic")
                except:
                    return self.draw_ascii_rebar(valid_segments)
            else:
                return self.draw_ascii_rebar(valid_segments)

    def save_figure_as_file(self, base64_data, filename, dpi=300):
        """
        將base64圖片數據儲存為檔案
        
        Args:
            base64_data: base64編碼的圖片數據
            filename: 儲存檔名
            dpi: 解析度（實際由原圖決定）
        
        Returns:
            str: 儲存的檔案路徑
        """
        try:
            # 解碼base64數據
            image_data = base64.b64decode(base64_data)
            
            # 確保目錄存在
            directory = os.path.dirname(filename)
            if directory and not os.path.exists(directory):
                os.makedirs(directory)
            
            # 儲存檔案
            with open(filename, 'wb') as f:
                f.write(image_data)
            
            return filename
        except Exception as e:
            raise Exception(f"圖檔儲存失敗: {str(e)}")

    @staticmethod
    def _figure_to_base64(fig):
        """將 matplotlib 圖形轉換為 base64 字串"""
        buf = io.BytesIO()
        fig.savefig(buf, format='png', bbox_inches='tight', 
                   pad_inches=0.1, dpi=200, facecolor='white', edgecolor='none')
        buf.seek(0)
        img_str = base64.b64encode(buf.read()).decode()
        plt.close(fig)
        return img_str


# 便利函數，讓使用更簡單
def create_graphics_manager():
    """創建圖形管理器實例"""
    return GraphicsManager()

def quick_draw_rebar(segments, rebar_number="#4", mode="professional"):
    """
    快速繪製鋼筋圖示的便利函數
    
    Args:
        segments: 分段長度列表
        rebar_number: 鋼筋編號
        mode: 繪圖模式
    
    Returns:
        tuple: (base64圖片數據, 詳細描述文字)
    """
    gm = GraphicsManager()
    
    # 檢查依賴
    deps_ok, missing = gm.check_dependencies()
    if not deps_ok:
        print(f"⚠️ 缺少套件: {missing}")
        mode = "ascii"
    
    # 生成圖示
    image_data = gm.generate_rebar_diagram(segments, rebar_number, mode)
    
    # 生成描述
    if mode != "ascii":
        description = gm.create_detailed_description(segments, rebar_number)
    else:
        description = f"ASCII 鋼筋圖示 {rebar_number}"
    
    return image_data, description


# 測試函數
def test_graphics_manager():
    """測試圖形管理器的各種功能"""
    gm = GraphicsManager()
    
    # 檢查依賴
    deps_ok, missing = gm.check_dependencies()
    print(f"圖形依賴檢查: {'✅ 通過' if deps_ok else '❌ 失敗'}")
    if missing:
        print(f"缺少套件: {missing}")
    
    # 測試案例
    test_cases = [
        ([300], "#4", "直鋼筋"),
        ([150, 200], "#5", "L型鋼筋"),
        ([120, 300, 120], "#6", "U型鋼筋"),
        ([100, 150, 200, 100], "#7", "複合鋼筋")
    ]
    
    print("\n🎨 圖形管理器測試")
    print("=" * 50)
    
    for i, (segments, rebar_num, desc) in enumerate(test_cases, 1):
        print(f"\n測試 {i}: {desc} {rebar_num}")
        print(f"分段: {segments}")
        
        try:
            # 測試專業模式
            if deps_ok:
                image_data = gm.generate_rebar_diagram(segments, rebar_num, "professional")
                description = gm.create_detailed_description(segments, rebar_num)
                
                # 儲存圖檔
                filename = f"test_rebar_{i}_{rebar_num.replace('#', '')}.png"
                gm.save_figure_as_file(image_data, filename)
                print(f"✅ 專業圖檔已儲存: {filename}")
                
                # 顯示部分描述
                lines = description.split('\n')
                print("📝 描述預覽:")
                for line in lines[:8]:  # 只顯示前8行
                    print(f"  {line}")
                print("  ...")
            else:
                # 測試ASCII模式
                ascii_art = gm.generate_rebar_diagram(segments, rebar_num, "ascii")
                print(f"📝 ASCII圖示: {ascii_art}")
                
        except Exception as e:
            print(f"❌ 錯誤: {str(e)}")
        
        print("-" * 30)


if __name__ == "__main__":
    """當檔案被直接執行時進行測試"""
    test_graphics_manager()