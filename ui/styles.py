"""
UI 樣式相關功能模組
"""

import tkinter as tk
from tkinter import ttk
from config import COLORS

class StyleManager:
    """UI 樣式管理器"""
    
    @staticmethod
    def setup_modern_styles():
        """設定現代化樣式"""
        style = ttk.Style()
        
        # 設定主題
        style.theme_use('clam')
        
        # Materialize 主色
        primary = COLORS['primary']
        secondary = COLORS['secondary']
        surface = COLORS['surface']
        background = COLORS['background']
        accent = COLORS['accent']
        
        # 設定按鈕樣式
        style.configure(
            'Modern.TButton',
            background=primary,
            foreground=COLORS['text_primary'],
            borderwidth=0,
            focusthickness=2,
            focuscolor=secondary,
            font=('Calibri', 11, 'bold'),
            relief='flat',
            padding=10
        )
        
        style.map(
            'Modern.TButton',
            background=[('active', secondary), ('disabled', accent)],
            foreground=[('disabled', COLORS['text_secondary'])]
        )
        
        # 設定標籤樣式
        style.configure(
            'Modern.TLabel',
            background=background,
            foreground=COLORS['text_primary'],
            font=('Calibri', 11)
        )
        
        # 設定標題標籤樣式
        style.configure(
            'Title.TLabel',
            background=background,
            foreground=primary,
            font=('Calibri', 16, 'bold')
        )
        
        # 設定輸入框樣式
        style.configure(
            'Modern.TEntry',
            fieldbackground=surface,
            foreground=COLORS['text_primary'],
            borderwidth=0,
            font=('Calibri', 11),
            relief='flat',
            padding=8
        )
        
        # 設定進度條樣式
        style.configure(
            'Modern.Horizontal.TProgressbar',
            background=primary,
            troughcolor=accent,
            borderwidth=0,
            thickness=8
        )
        
        # 設定框架樣式
        style.configure(
            'Card.TFrame',
            background=surface,
            relief='flat',
            borderwidth=0,
            padding=18
        )
        
        # 設定選單樣式
        style.configure(
            'Modern.TMenubutton',
            background=COLORS['surface'],
            foreground=COLORS['text_primary'],
            borderwidth=1,
            focusthickness=1,
            focuscolor=COLORS['primary'],
            font=('微軟正黑體', 10)
        )
        
        style.map(
            'Modern.TMenubutton',
            background=[('active', COLORS['accent']), ('disabled', COLORS['accent'])],
            foreground=[('disabled', COLORS['text_secondary'])]
        )

        # 次要按鈕樣式
        style.configure(
            'Secondary.TButton',
            background=accent,
            foreground=COLORS['text_primary'],
            borderwidth=0,
            focusthickness=2,
            focuscolor=secondary,
            font=('Calibri', 11, 'bold'),
            relief='flat',
            padding=10
        )
        style.map(
            'Secondary.TButton',
            background=[('active', surface), ('disabled', accent)],
            foreground=[('disabled', COLORS['text_secondary'])]
        )

class ModernButton(ttk.Button):
    """現代化按鈕元件（ttk 版本，套用 Modern.TButton 樣式）"""
    def __init__(self, master, **kwargs):
        kwargs.setdefault('style', 'Modern.TButton')
        super().__init__(master, **kwargs)

class CardFrame(ttk.Frame):
    """卡片式框架"""
    def __init__(self, master, title=None, padding=24, **kwargs):
        super().__init__(master, style='Card.TFrame', **kwargs)
        self.padding = padding
        self.configure(padding=padding)
        self['borderwidth'] = 0
        self['relief'] = 'flat'
        if title:
            self.title_label = tk.Label(
                self,
                text=title,
                bg=COLORS['surface'],
                fg=COLORS['primary'],
                font=('Calibri', 15, 'bold'),
                pady=10
            )
            self.title_label.pack(pady=(8, 8))

    def configure(self, *args, **kwargs):
        if 'padding' in kwargs:
            padding = kwargs.pop('padding')
            kwargs['padding'] = (padding, padding, padding, padding)
        super().configure(*args, **kwargs)

class ModernEntry(ttk.Entry):
    """現代化輸入框元件"""
    
    def __init__(self, master, **kwargs):
        kwargs.setdefault('style', 'Modern.TEntry')
        super().__init__(master, **kwargs)

class ModernLabel(ttk.Label):
    """現代化標籤元件"""
    
    def __init__(self, master, **kwargs):
        kwargs.setdefault('style', 'Modern.TLabel')
        super().__init__(master, **kwargs) 