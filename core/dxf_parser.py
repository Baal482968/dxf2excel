import re

class DxfParser:
    # ... existing code ...
    def _parse_rebar_mark(self, text: str):
        """
        解析鋼筋標記，支援多種格式（含前綴、長度小數點、單段、多段）
        """
        # 移除前綴非 # 的部分，但保留 N# 這種情形
        m = re.search(r'((N#|#)[0-9]+-.*?x[0-9]+)', text, re.IGNORECASE)
        if m:
            text = m.group(1)
        # 例：N#10-1000+1000+1000x31、#10-510.5x11、#9-45+700x50
        m = re.match(r'(N#|#)(\d+)-([\d\.]+(?:\+[\d\.]+)*)x(\d+)', text, re.IGNORECASE)
        if m:
            rebar_number = f"{m.group(1)}{m.group(2)}"  # 保留 N# 或 #
            segments = [float(x) if '.' in x else int(x) for x in m.group(3).split('+')]
            count = int(m.group(4))
            return {
                'rebar_number': rebar_number,
                'segments': segments,
                'count': count
            }
        return None 