"""
鋼筋處理相關功能模組
"""

import re
from config import REBAR_UNIT_WEIGHT, REBAR_DIAMETERS, REBAR_GRADES
# 圖形相關模組已移除，改為使用 assets/materials/ 資料夾中的圖示檔案

class RebarProcessor:
    """鋼筋處理器"""
    
    @staticmethod
    def extract_rebar_info(text):
        """
        從文字中提取鋼筋資訊
        
        格式範例：
        - #4@20
        - #4@20c/c
        - #4@20cm
        - #4@200mm
        """
        # 移除空白字元
        text = text.strip()
        
        # 基本格式檢查
        if not text or not text.startswith('#'):
            return None
        
        try:
            # 提取鋼筋編號
            rebar_number = re.match(r'#\d+', text).group()
            
            # 提取間距
            spacing_match = re.search(r'@(\d+)(?:c/c|cm|mm)?', text)
            if spacing_match:
                spacing = int(spacing_match.group(1))
                # 如果沒有單位，預設為公分
                if not re.search(r'@\d+(?:c/c|cm|mm)', text):
                    spacing *= 10  # 轉換為公釐
            else:
                spacing = None
            
            return {
                'rebar_number': rebar_number,
                'spacing': spacing,
                'diameter': RebarProcessor.get_rebar_diameter(rebar_number),
                'unit_weight': RebarProcessor.get_rebar_unit_weight(rebar_number),
                'grade': RebarProcessor.get_rebar_grade(rebar_number)
            }
        except Exception:
            return None

    @staticmethod
    def get_rebar_diameter(number):
        """獲取鋼筋直徑（公釐）"""
        return REBAR_DIAMETERS.get(number, 0)

    @staticmethod
    def get_rebar_unit_weight(number):
        """獲取鋼筋單位重量（kg/m）"""
        return REBAR_UNIT_WEIGHT.get(number, 0)

    @staticmethod
    def get_rebar_grade(number):
        """獲取鋼筋材質等級"""
        return REBAR_GRADES.get(number, "未知")

    @staticmethod
    def calculate_rebar_weight(number, length, count=1):
        """計算鋼筋重量（kg）"""
        unit_weight = RebarProcessor.get_rebar_unit_weight(number)
        return unit_weight * length * count

    @staticmethod
    def parse_rebar_text(text):
        """
        支援多種鋼筋標記格式的解析，允許 # 前有任意字元，支援全形乘號、逗號、cm、括號等符號，並自動解析角度。
        
        支援格式：
        1. #10-1000+1000+1000x31
        2. #10-21+550x6
        3. #10-1400x1
        4. 折140#10-1000+1200x20
        5. #10-510.5x11
        6. #10-550+21x5
        7. 地箍#5(50x75)=20
        8. U箍#5(50x75)=20
        9. L箍#4(50+75)=20
        10. 柱繫#5-3000x20
        11. 梁繫#5-3000x20
        12. 牆繫#5-3000x20
        13. 安#5-3000x20
        14. 安#5-25+1000x20
        """
        import re
        text = text.strip()
        
        # 新增：處理箍筋格式
        # 格式: (地箍|U箍|柱箍|牆箍|L箍|半箍)#5(50x75)=20 或 (50+75)=20
        stirrup_pattern = r'(地箍|U箍|柱箍|牆箍|L箍|半箍)(#\d+)\((\d+(?:\.\d+)?)\s?([x+])\s?(\d+(?:\.\d+)?)\)=(\d+)'
        stirrup_match = re.match(stirrup_pattern, text)

        if stirrup_match:
            stirrup_type = stirrup_match.group(1)
            rebar_number = stirrup_match.group(2)
            width = float(stirrup_match.group(3))
            operator = stirrup_match.group(4)
            height = float(stirrup_match.group(5))
            count = int(stirrup_match.group(6))
            
            # 根據不同箍筋類型計算長度
            # 這裡的長度計算是根據您提供的圖片範例推斷的近似值
            # 實際長度可能需要更複雜的彎鉤長度計算標準
            length = 0
            if stirrup_type in ['地箍', '柱箍', '牆箍']: # 封閉箍筋
                length = (width + height) * 2
                if rebar_number == '#5':
                    length += 30 # 根據圖例推斷的彎鉤長度
                elif rebar_number == '#4':
                    length += 24 # 根據圖例推斷的彎鉤長度
            elif stirrup_type == 'U箍':
                length = width + height * 2
                if rebar_number == '#5':
                    length += 30 # 根據圖例推斷的彎鉤長度
            elif stirrup_type == 'L箍':
                # L箍和半箍的(寬+高)格式代表的是總長，此處為解析出的片段
                length = width + height
                if rebar_number == '#4':
                     # 根據圖例，長度為 149，(50+75)=125，差24
                     # 這24可能是彎鉤或額外長度
                    length += 24
            elif stirrup_type == '半箍':
                length = width + height
                if rebar_number == '#3':
                    # 根據圖例，長度為 145，(50+75)=125，差20
                    length += 20

            return {
                'rebar_number': rebar_number,
                'segments': [width, height],
                'angles': [], # 角度資訊目前無法從文字中解析
                'count': count,
                'raw_text': text,
                'length': length,
                'type': stirrup_type
            }
        
        # 新增：處理繫筋格式
        # 格式: (柱繫|梁繫|牆繫)#5-3000x20
        tie_pattern = r'(柱繫|梁繫|牆繫)(#\d+)-([\d\.]+?)x(\d+)'
        tie_match = re.match(tie_pattern, text)

        if tie_match:
            tie_type = tie_match.group(1)
            rebar_number = tie_match.group(2)
            length = float(tie_match.group(3))
            count = int(tie_match.group(4))

            return {
                'rebar_number': rebar_number,
                'segments': [length],
                'angles': [],
                'count': count,
                'raw_text': text,
                'length': length,
                'type': tie_type,
                'note': tie_type # 在備註中也標明類型
            }
        # 新增：處理 type10 直料鋼筋格式
        # 格式: #3-700x99 (單段直料)
        type10_pattern = r'(#\d+)-([\d\.]+)x(\d+)'
        type10_match = re.match(type10_pattern, text)

        if type10_match:
            rebar_number = type10_match.group(1)
            length = float(type10_match.group(2))
            count = int(type10_match.group(3))
            
            # 計算重量
            unit_weight = RebarProcessor.get_rebar_unit_weight(rebar_number)
            weight = unit_weight * length * count / 100  # 轉換為 kg

            return {
                'rebar_number': rebar_number,
                'segments': [length],
                'angles': [],
                'count': count,
                'raw_text': text,
                'length': length,
                'weight': weight,
                'type': 'type10',
                'note': '直料'
            }
        
        # 新增：處理「安」筋格式
        # 格式: 安#5-3000x20 or 安#5-25+1000x20
        anchor_pattern = r'(安)(#\d+)-([\d\.]+(?:\+[\d\.]+)*)x(\d+)'
        anchor_match = re.match(anchor_pattern, text)

        if anchor_match:
            anchor_type = anchor_match.group(1)
            rebar_number = anchor_match.group(2)
            segments_str = anchor_match.group(3)
            count = int(anchor_match.group(4))
            segments = [float(x) if '.' in x else int(x) for x in segments_str.split('+')]

            return {
                'rebar_number': rebar_number,
                'segments': segments,
                'angles': [],
                'count': count,
                'raw_text': text,
                'length': sum(segments),
                'type': anchor_type,
                'note': anchor_type
            }

        # Debug: 印出原始文字
        # print(f"[DEBUG] 解析鋼筋文字: {text}")
        
        # 解析角度（如「折135度」「135°」）
        angles = []
        angle_match = re.search(r'(?:折)?(\d{2,3})[°度]', text)
        if angle_match:
            try:
                angle = int(angle_match.group(1))
                angles.append(angle)
                # print(f"[DEBUG] 解析到角度: {angle}°")
            except:
                pass
        
        # 將全形乘號、逗號、空白等轉成半形
        text = text.replace('×', 'x').replace('，', ',').replace(' ', '')
        
        # 支援折開頭格式
        if text.startswith('折'):
            # 解析號數
            m = re.search(r'#(\d+)', text)
            rebar_no = f"#{m.group(1)}" if m else ""
            # 彎曲鋼筋解析功能已移除，暫時跳過處理
            pass
        
        # 強化正則：支援 N# 或 # 開頭
        pattern = r'(N#|#)(\d+)-([\d\.]+(?:\+[\d\.]+)*)x(\d+)'
        m = re.match(pattern, text, re.IGNORECASE)
        if m:
            number = f"{m.group(1)}{m.group(2)}"  # 保留 N# 或 #
            segments_str = m.group(3)
            count = int(m.group(4))
            # 解析分段
            segments = []
            for seg in segments_str.split('+'):
                try:
                    if '.' in seg:
                        segments.append(float(seg))
                    else:
                        segments.append(int(seg))
                except ValueError:
                    continue
            # 確保 segments 不為空
            if not segments:
                return None
            result = {
                'rebar_number': number,
                'segments': segments,
                'angles': angles,
                'count': count,
                'raw_text': text,
                'length': sum(segments)
            }
            return result
            
        # print(f"[DEBUG] 無法解析鋼筋文字格式: {text}")
        return None

    @staticmethod
    def validate_rebar_number(number):
        """驗證鋼筋編號是否有效"""
        return number in REBAR_UNIT_WEIGHT

    @staticmethod
    def get_rebar_summary(rebar_list):
        """生成鋼筋統計摘要"""
        summary = {}
        
        for rebar in rebar_list:
            number = rebar['rebar_number']
            if number not in summary:
                summary[number] = {
                    'count': 0,
                    'total_length': 0,
                    'total_weight': 0,
                    'diameter': RebarProcessor.get_rebar_diameter(number),
                    'grade': RebarProcessor.get_rebar_grade(number)
                }
            
            summary[number]['count'] += 1
            summary[number]['total_length'] += rebar.get('length', 0)
            summary[number]['total_weight'] += rebar.get('weight', 0)
        
        return summary 