"""
繪製U形鋼筋
"""
import matplotlib.pyplot as plt
from .common import figure_to_base64, draw_dimension_line

def draw_u_shaped_rebar(length1, length2, length3, rebar_number, professional=True, width=700, height=260, settings=None, complex_drawer=None):
    """繪製 U 型或階梯型鋼筋圖示（自動判斷對稱或階梯）"""
    # 若兩側長度相等，畫對稱U型，否則畫階梯型
    if abs(length1 - length3) < 1e-3:
        if professional:
            return _draw_professional_u_shaped_rebar(length1, length2, length3, rebar_number, width, height)
        else:
            if settings is None:
                raise ValueError("Basic settings must be provided for basic mode")
            return _draw_basic_u_shaped_rebar(length1, length2, length3, rebar_number, width, height, settings)
    else:
        # 階梯型視為複雜型
        if complex_drawer:
            return complex_drawer([length1, length2, length3], rebar_number, professional, angles=None, width=width, height=height, settings=settings)
        else:
            # Fallback: 使用預設的 U 型繪製
            print("[WARNING] complex_drawer not provided for non-symmetric U-shape, using default U-shape drawing.")
            return _draw_professional_u_shaped_rebar(length1, length2, length3, rebar_number, width, height)


def _draw_professional_u_shaped_rebar(length1, length2, length3, rebar_number, width=700, height=260):
    """極簡U型鋼筋圖示（橫線加長，標註分散，不顯示編號）"""
    fig, ax = plt.subplots(figsize=(width/100, height/100))
    ax.set_aspect('equal')
    # 預留邊界
    margin_x = 18
    margin_y = 16
    # 橫線加長倍率
    hor_scale = 4.0
    # 計算縮放比例
    l1 = float(length1)
    l2 = float(length2) * hor_scale
    l3 = float(length3)
    total_w = l2
    total_h = l1 + l3
    scale_x = (width - 2*margin_x) / total_w if total_w > 0 else 1
    scale_y = (height - 2*margin_y) / total_h if total_h > 0 else 1
    scale = min(scale_x, scale_y)
    # 重新計算各段長度
    l1s = float(length1) * scale
    l2s = float(length2) * hor_scale * scale
    l3s = float(length3) * scale
    # U 字形：橫線在下，兩豎線朝下
    # 起點設在左上角
    p1 = (margin_x, margin_y)  # 左上
    p2 = (margin_x, margin_y + l1s)  # 左下
    p3 = (margin_x + l2s, margin_y + l1s)  # 右下
    p4 = (margin_x + l2s, margin_y)  # 右上
    # 主線（黑色，細線）
    ax.plot([p1[0], p2[0]], [p1[1], p2[1]], color='black', linewidth=1.2, solid_capstyle='butt')  # 左豎
    ax.plot([p2[0], p3[0]], [p2[1], p3[1]], color='black', linewidth=1.2, solid_capstyle='butt')  # 橫（下）
    ax.plot([p3[0], p4[0]], [p3[1], p4[1]], color='black', linewidth=1.2, solid_capstyle='butt')  # 右豎
    # 長度標註（左、下、右，黑色，無粗體，字體小）
    ax.text(p1[0]-10, (p1[1]+p2[1])/2, f'{int(length1)}', ha='right', va='center', fontsize=13, color='black')
    ax.text((p2[0]+p3[0])/2, p2[1]-10, f'{int(length2)}', ha='center', va='bottom', fontsize=13, color='black')
    ax.text(p4[0]+10, (p4[1]+p3[1])/2, f'{int(length3)}', ha='left', va='center', fontsize=13, color='black')
    # 不顯示編號
    ax.set_xlim(0, width)
    ax.set_ylim(height, 0)  # y軸反轉，讓原點在左上
    ax.axis('off')
    plt.tight_layout(pad=0.2)
    return figure_to_base64(fig)

def _draw_basic_u_shaped_rebar(length1, length2, length3, rebar_number, width=700, height=260, settings=None):
    """繪製基本U型鋼筋圖示（保留原有功能）"""
    fig, ax = plt.subplots(figsize=(width/300, height/300))
    
    # 繪製 U 型鋼筋
    ax.plot([0, length1], [0, 0], 
            color=settings['colors']['rebar'], 
            linewidth=settings['line_width'])
    ax.plot([length1, length1], [0, -length2], 
            color=settings['colors']['rebar'], 
            linewidth=settings['line_width'])
    ax.plot([length1, 0], [-length2, -length2], 
            color=settings['colors']['rebar'], 
            linewidth=settings['line_width'])
    ax.plot([0, 0], [-length2, 0], 
            color=settings['colors']['rebar'], 
            linewidth=settings['line_width'])
    
    ax.set_xlim(-1, length1 + 1)
    ax.set_ylim(-length2 - 1, 1)
    ax.axis('off')
    
    # 加入鋼筋編號標記
    ax.text(length1/2, 0.3, rebar_number, ha='center', va='center')
    
    return figure_to_base64(fig) 