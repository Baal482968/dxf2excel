"""
此目錄包含各種鋼筋形狀的繪製邏輯。
每個檔案都應包含一個 `draw` 函數，該函數接收 `ax` (matplotlib軸)、`segments`、`rebar_number` 和 `settings`。
"""
from .straight import draw_straight_rebar
from .l_shape import draw_l_shaped_rebar
from .u_shape import draw_u_shaped_rebar
from .n_shape import draw_n_shaped_rebar
from .bent import draw_bent_rebar, parse_bent_rebar_string
from .complex_shape import draw_complex_rebar 