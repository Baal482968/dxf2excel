"""
繪製N形鋼筋
"""
import matplotlib.pyplot as plt
from .common import figure_to_base64

def draw_n_shaped_rebar(length1, length2, length3, rebar_number, width=700, height=260, settings=None):
    """繪製 N 型鋼筋圖示（左上→左中，左中→右中，右中→右下，兩豎線都朝下，橫線水平）"""
    fig, ax = plt.subplots(figsize=(width/100, height/100))
    ax.set_aspect('equal')
    # 固定長度
    ver_len = 32   # 豎線長度(px)
    hor_len = 80   # 橫線長度(px)
    margin_x = 30
    margin_y = 20
    # 座標
    p1 = (margin_x, margin_y)  # 左上
    p2 = (margin_x, margin_y + ver_len)  # 左中
    p3 = (margin_x + hor_len, margin_y + ver_len)  # 右中
    p4 = (margin_x + hor_len, margin_y + ver_len + ver_len)  # 右下
    # 主線
    ax.plot([p1[0], p2[0]], [p1[1], p2[1]], color='black', linewidth=1.2)  # 左豎
    ax.plot([p2[0], p3[0]], [p2[1], p3[1]], color='black', linewidth=1.2)  # 水平橫線
    ax.plot([p3[0], p4[0]], [p3[1], p4[1]], color='black', linewidth=1.2)  # 右豎
    # 長度標註
    ax.text(p1[0]-10, (p1[1]+p2[1])/2, f'{int(length1)}', ha='right', va='center', fontsize=13, color='black')
    ax.text((p2[0]+p3[0])/2, p2[1]-10, f'{int(length2)}', ha='center', va='top', fontsize=13, color='black')
    ax.text(p4[0]+10, (p3[1]+p4[1])/2, f'{int(length3)}', ha='left', va='center', fontsize=13, color='black')
    ax.set_xlim(0, width)
    ax.set_ylim(0, height)
    ax.axis('off')
    plt.tight_layout(pad=0.2)
    return figure_to_base64(fig) 