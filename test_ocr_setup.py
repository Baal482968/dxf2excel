#!/usr/bin/env python3
"""
測試 OCR 環境設置
"""

import sys
import os

def test_imports():
    """測試必要的套件是否已安裝"""
    print("🔍 測試套件導入...")
    
    try:
        import cv2
        print("✅ OpenCV 已安裝")
    except ImportError:
        print("❌ OpenCV 未安裝")
        return False
    
    try:
        import numpy as np
        print("✅ NumPy 已安裝")
    except ImportError:
        print("❌ NumPy 未安裝")
        return False
    
    try:
        import PIL
        print("✅ Pillow 已安裝")
    except ImportError:
        print("❌ Pillow 未安裝")
        return False
    
    try:
        import pytesseract
        print("✅ pytesseract 已安裝")
    except ImportError:
        print("❌ pytesseract 未安裝")
        return False
    
    return True

def test_tesseract():
    """測試 Tesseract 是否可用"""
    print("\n🔍 測試 Tesseract OCR 引擎...")
    
    try:
        import pytesseract
        
        # Windows 環境下的路徑設定
        if sys.platform.startswith('win'):
            possible_paths = [
                r'C:\Program Files\Tesseract-OCR\tesseract.exe',
                r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
                r'C:\Users\{}\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'.format(os.getenv('USERNAME', '')),
            ]
            
            for path in possible_paths:
                if os.path.exists(path):
                    pytesseract.pytesseract.tesseract_cmd = path
                    print(f"✅ 找到 Tesseract: {path}")
                    break
            else:
                print("⚠️  未找到 Tesseract，嘗試使用 PATH 中的版本")
        
        # 檢查 Tesseract 版本
        version = pytesseract.get_tesseract_version()
        print(f"✅ Tesseract 版本: {version}")
        
        # 檢查可用的語言
        languages = pytesseract.get_languages()
        print(f"✅ 可用語言: {languages}")
        
        return True
        
    except Exception as e:
        print(f"❌ Tesseract 測試失敗: {str(e)}")
        print("\n📋 請安裝 Tesseract:")
        if sys.platform.startswith('win'):
            print("   Windows 方法:")
            print("   1. 下載: https://github.com/UB-Mannheim/tesseract/wiki")
            print("   2. 安裝到: C:\\Program Files\\Tesseract-OCR")
            print("   3. 設定環境變數 PATH")
            print("   4. 或使用: choco install tesseract")
        else:
            print("   macOS/Linux 方法:")
            print("   1. brew install tesseract")
            print("   2. sudo apt-get install tesseract-ocr")
            print("   3. sudo port install tesseract")
        return False

def test_image_ocr():
    """測試圖片 OCR 功能"""
    print("\n🔍 測試圖片 OCR 功能...")
    
    try:
        import cv2
        import numpy as np
        import pytesseract
        
        # 創建一個簡單的測試圖片（包含數字）
        test_image = np.ones((100, 200), dtype=np.uint8) * 255
        cv2.putText(test_image, "123", (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 2, 0, 3)
        
        # 嘗試 OCR
        result = pytesseract.image_to_string(test_image, config='--psm 6')
        print(f"✅ OCR 測試成功，識別結果: '{result.strip()}'")
        
        return True
        
    except Exception as e:
        print(f"❌ OCR 測試失敗: {str(e)}")
        return False

def main():
    """主函數"""
    print("🚀 OCR 環境測試")
    print("=" * 50)
    print(f"作業系統: {sys.platform}")
    print()
    
    # 測試套件導入
    if not test_imports():
        print("\n❌ 套件導入失敗，請先安裝必要套件")
        return
    
    # 測試 Tesseract
    if not test_tesseract():
        print("\n❌ Tesseract 未正確安裝")
        return
    
    # 測試 OCR 功能
    if not test_image_ocr():
        print("\n❌ OCR 功能測試失敗")
        return
    
    print("\n🎉 所有測試通過！OCR 環境已正確設置")
    print("現在可以測試 DXF 圖片 OCR 功能了")

if __name__ == "__main__":
    main() 