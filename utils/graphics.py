"""
圖形繪製相關功能模組
"""

import matplotlib.pyplot as plt
import numpy as np
from PIL import Image
import io
import base64

class GraphicsManager:
    """圖形繪製管理器"""
    
    @staticmethod
    def check_dependencies():
        """檢查圖形繪製所需的套件是否已安裝"""
        try:
            import matplotlib
            import numpy
            from PIL import Image
            return True
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

    @staticmethod
    def draw_straight_rebar(length, rebar_number):
        """繪製直鋼筋圖示"""
        fig, ax = plt.subplots(figsize=(6, 1))
        ax.plot([0, length], [0, 0], 'b-', linewidth=2)
        ax.set_xlim(-1, length + 1)
        ax.set_ylim(-1, 1)
        ax.axis('off')
        
        # 加入鋼筋編號標記
        ax.text(length/2, 0.3, rebar_number, ha='center', va='center')
        
        return GraphicsManager._figure_to_base64(fig)

    @staticmethod
    def draw_l_shaped_rebar(length1, length2, rebar_number):
        """繪製 L 型鋼筋圖示"""
        fig, ax = plt.subplots(figsize=(6, 3))
        
        # 繪製 L 型鋼筋
        ax.plot([0, length1], [0, 0], 'b-', linewidth=2)
        ax.plot([length1, length1], [0, -length2], 'b-', linewidth=2)
        
        ax.set_xlim(-1, length1 + 1)
        ax.set_ylim(-length2 - 1, 1)
        ax.axis('off')
        
        # 加入鋼筋編號標記
        ax.text(length1/2, 0.3, rebar_number, ha='center', va='center')
        
        return GraphicsManager._figure_to_base64(fig)

    @staticmethod
    def draw_u_shaped_rebar(length1, length2, length3, rebar_number):
        """繪製 U 型鋼筋圖示"""
        fig, ax = plt.subplots(figsize=(6, 3))
        
        # 繪製 U 型鋼筋
        ax.plot([0, length1], [0, 0], 'b-', linewidth=2)
        ax.plot([length1, length1], [0, -length2], 'b-', linewidth=2)
        ax.plot([length1, 0], [-length2, -length2], 'b-', linewidth=2)
        ax.plot([0, 0], [-length2, 0], 'b-', linewidth=2)
        
        ax.set_xlim(-1, length1 + 1)
        ax.set_ylim(-length2 - 1, 1)
        ax.axis('off')
        
        # 加入鋼筋編號標記
        ax.text(length1/2, 0.3, rebar_number, ha='center', va='center')
        
        return GraphicsManager._figure_to_base64(fig)

    @staticmethod
    def draw_complex_rebar(segments, rebar_number):
        """繪製複雜多段鋼筋圖示"""
        fig, ax = plt.subplots(figsize=(8, 4))
        
        # 計算總長度和高度
        total_length = sum(segments)
        max_height = max(segments)
        
        # 繪製多段鋼筋
        x, y = 0, 0
        for i, length in enumerate(segments):
            if i % 2 == 0:
                ax.plot([x, x + length], [y, y], 'b-', linewidth=2)
                x += length
            else:
                ax.plot([x, x], [y, y - length], 'b-', linewidth=2)
                y -= length
        
        ax.set_xlim(-1, total_length + 1)
        ax.set_ylim(-max_height - 1, 1)
        ax.axis('off')
        
        # 加入鋼筋編號標記
        ax.text(total_length/2, 0.3, rebar_number, ha='center', va='center')
        
        return GraphicsManager._figure_to_base64(fig)

    @staticmethod
    def draw_ascii_rebar(segments):
        """生成 ASCII 格式的鋼筋圖示"""
        if not segments:
            return ""
            
        # 計算圖示寬度
        width = sum(segments)
        height = max(segments)
        
        # 創建空白畫布
        canvas = [[' ' for _ in range(width)] for _ in range(height + 2)]
        
        # 繪製鋼筋
        x, y = 0, 0
        for i, length in enumerate(segments):
            if i % 2 == 0:
                # 水平線
                for j in range(length):
                    canvas[y][x + j] = '-'
                x += length
            else:
                # 垂直線
                for j in range(length):
                    canvas[y - j][x] = '|'
                y -= length
        
        # 轉換為字串
        return '\n'.join([''.join(row) for row in canvas])

    @staticmethod
    def _figure_to_base64(fig):
        """將 matplotlib 圖形轉換為 base64 字串"""
        buf = io.BytesIO()
        fig.savefig(buf, format='png', bbox_inches='tight', pad_inches=0.1)
        buf.seek(0)
        img_str = base64.b64encode(buf.read()).decode()
        plt.close(fig)
        return img_str 