"""
繪製折彎鋼筋
"""
import matplotlib.pyplot as plt
import numpy as np
import re
from .common import figure_to_base64

def draw_bent_rebar(angle, length1, length2, rebar_number, width=700, height=260, settings=None):
    """繪製折彎鋼筋圖示（工程圖標準：左水平→折點→右水平，標註位置正確，置中顯示）"""
    fig, ax = plt.subplots(figsize=(width/100, height/100))
    ax.set_aspect('equal')
    # 固定顯示長度
    L1 = 120
    L2 = 120
    margin = 70
    # 起點設在左側中間
    x0, y0 = margin, height/2
    # 第一段：水平線
    x1, y1 = x0 + L1, y0
    # 第二段：折點後，逆時針 angle
    theta = np.radians(180 - angle)  # 工程圖標準，逆時針
    x2 = x1 + L2 * np.cos(theta)
    y2 = y1 + L2 * np.sin(theta)
    # 畫線
    ax.plot([x0, x1], [y0, y1], color='black', linewidth=2.2)
    ax.plot([x1, x2], [y1, y2], color='black', linewidth=2.2)
    # 長度標註
    ax.text((x0+x1)/2 - 30, y0-38, f'{int(length1)}', ha='center', va='center', fontsize=12)
    ax.text((x1+x2)/2 + 10, (y1+y2)/2, f'{int(length2)}', ha='left', va='center', fontsize=12)
    ax.text(x1-18, y1+28, f'{int(angle)}°', ha='center', va='center', fontsize=12)
    ax.axis('off')
    return figure_to_base64(fig)

def parse_bent_rebar_string(rebar_number_str):
    """解析折彎鋼筋字串，回傳 (角度, 長度1, 長度2)"""
    # 支援格式：折140#10-1000+1200 或 折140#10-1000+1200x20
    # 先抓角度
    m = re.match(r'^折(\d+)', rebar_number_str)
    angle = int(m.group(1)) if m else 0
    # 再抓號數後的長度們
    m2 = re.search(r'#\d+-([\d\.]+)\+([\d\.]+)', rebar_number_str)
    if m2:
        length1 = int(float(m2.group(1)))
        length2 = int(float(m2.group(2)))
        return angle, length1, length2
    return None 