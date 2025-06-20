"""
繪製直鋼筋
"""
import matplotlib.pyplot as plt
from .common import figure_to_base64

def draw_straight_rebar(length, rebar_number, professional=True, width=700, height=260, settings=None):
    """繪製直鋼筋圖示"""
    if professional:
        return _draw_professional_straight_rebar(length, rebar_number, width, height)
    else:
        if settings is None:
            raise ValueError("Basic settings must be provided for basic mode")
        return _draw_basic_straight_rebar(length, rebar_number, width, height, settings)

def _draw_professional_straight_rebar(length, rebar_number, width=700, height=260):
    """極簡直鋼筋圖示（放大線條與字體，內容置中）"""
    fig, ax = plt.subplots(figsize=(width/100, height/100))
    ax.set_aspect('equal')
    # 動態縮放比例，讓線條最長佔畫布 90%，最短佔 30%
    min_ratio = 0.3
    max_ratio = 0.9
    min_length = 100
    max_length = 2000
    # 根據 length 線性插值
    ratio = min_ratio + (max_ratio - min_ratio) * (min(max(length, min_length), max_length) - min_length) / (max_length - min_length)
    length_scaled = width * ratio
    # 置中計算
    start_x = (width - length_scaled) / 2
    start_y = height / 2
    end_x = start_x + length_scaled
    # 主線
    ax.plot([start_x, end_x], [start_y, start_y], color='#2C3E50', linewidth=4, solid_capstyle='round')
    # 長度標註（線下方，黑色，粗體，白色底框，明顯在線下方）
    ax.text((start_x+end_x)/2, start_y+36, f'{int(length)}', ha='center', va='bottom', fontsize=32, color='black', fontweight='bold', bbox=dict(boxstyle="square,pad=0.2", facecolor='white', edgecolor='none', alpha=1.0))
    # 不顯示編號
    ax.set_xlim(0, width)
    ax.set_ylim(0, height)
    ax.axis('off')
    plt.tight_layout(pad=0.2)
    return figure_to_base64(fig)

def _draw_basic_straight_rebar(length, rebar_number, width=700, height=260, settings=None):
    """繪製基本直鋼筋圖示（保留原有功能）"""
    fig, ax = plt.subplots(figsize=(width/300, height/300))
    ax.plot([0, length], [0, 0], 
            color=settings['colors']['rebar'], 
            linewidth=settings['line_width'])
    ax.set_xlim(-1, length + 1)
    ax.set_ylim(-1, 1)
    ax.axis('off')
    
    # 加入鋼筋編號標記
    ax.text(length/2, 0.3, rebar_number, ha='center', va='center')
    
    return figure_to_base64(fig) 