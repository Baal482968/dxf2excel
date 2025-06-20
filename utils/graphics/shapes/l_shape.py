"""
繪製L形鋼筋
"""
import matplotlib.pyplot as plt
from .common import figure_to_base64

def draw_l_shaped_rebar(length1, length2, rebar_number, professional=True, width=700, height=260, settings=None):
    """繪製 L 型鋼筋圖示（線段長度固定，標註數字依 DXF 文字順序）"""
    if professional:
        return _draw_professional_l_shaped_rebar(length1, length2, rebar_number, width, height)
    else:
        if settings is None:
            raise ValueError("Basic settings must be provided for basic mode")
        return _draw_basic_l_shaped_rebar(length1, length2, rebar_number, width, height, settings)

def _draw_professional_l_shaped_rebar(length1, length2, rebar_number, width=700, height=260):
    """極簡L型鋼筋圖示，線段長度固定，標註數字依 DXF 文字順序"""
    fig, ax = plt.subplots(figsize=(width/100, height/100))
    ax.set_aspect('equal')
    # 固定圖形長度
    hor_len = 220  # 橫線長度(px)
    ver_len = 80   # 直線長度(px)
    center_x = width / 2
    center_y = height / 2
    # 左側直線
    start_x = center_x - hor_len / 2
    start_y = center_y + ver_len / 2
    # 畫直線（左側）
    ax.plot([start_x, start_x], [start_y, start_y - ver_len], color='#2C3E50', linewidth=5, solid_capstyle='round')
    # 畫橫線（下方）
    ax.plot([start_x, start_x + hor_len], [start_y - ver_len, start_y - ver_len], color='#2C3E50', linewidth=5, solid_capstyle='round')
    # segments[0] 標在直線中央
    ax.text(start_x - 20, start_y - ver_len/2, f'{int(length1)}', ha='right', va='center', fontsize=32, color='black', fontweight='bold', bbox=dict(boxstyle="square,pad=0.3", facecolor='white', edgecolor='none', alpha=1.0))
    # segments[1] 標在橫線下方中央
    ax.text(start_x + hor_len/2, start_y - ver_len - 20, f'{int(length2)}', ha='center', va='top', fontsize=28, color='black', fontweight='bold', bbox=dict(boxstyle="square,pad=0.3", facecolor='white', edgecolor='none', alpha=1.0))
    ax.set_xlim(0, width)
    ax.set_ylim(0, height)
    ax.axis('off')
    plt.tight_layout(pad=0.2)
    return figure_to_base64(fig)

def _draw_basic_l_shaped_rebar(length1, length2, rebar_number, width=700, height=260, settings=None):
    """繪製基本L型鋼筋圖示（保留原有功能）"""
    fig, ax = plt.subplots(figsize=(width/300, height/300))
    
    # 繪製 L 型鋼筋
    ax.plot([0, length1], [0, 0], 
            color=settings['colors']['rebar'], 
            linewidth=settings['line_width'])
    ax.plot([length1, length1], [0, -length2], 
            color=settings['colors']['rebar'], 
            linewidth=settings['line_width'])
    
    ax.set_xlim(-1, length1 + 1)
    ax.set_ylim(-length2 - 1, 1)
    ax.axis('off')
    
    # 加入鋼筋編號標記
    ax.text(length1/2, 0.3, rebar_number, ha='center', va='center')
    
    return figure_to_base64(fig) 