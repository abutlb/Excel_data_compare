#!/usr/bin/env python
# -*- coding: utf-8 -*-

import tkinter as tk
from excel_comparer_app import ExcelComparerApp

def main():
    """نقطة البداية للبرنامج"""
    root = tk.Tk()
    root.title("برنامج مقارنة ملفات الإكسل")
    
    # تعيين الخط الافتراضي للتطبيق (يدعم اللغة العربية)
    try:
        root.option_add("*Font", "Arial 10")
    except:
        pass
    
    # تهيئة التطبيق
    app = ExcelComparerApp(root)
    
    # تشغيل الحلقة الرئيسية
    root.mainloop()

if __name__ == "__main__":
    main()