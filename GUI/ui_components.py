#!/usr/bin/env python
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, scrolledtext

def create_ui(app):
    """إنشاء واجهة المستخدم الرسومية"""
    # الإطار الرئيسي
    main_frame = ttk.Frame(app.root, padding="10")
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # إطار اختيار الملفات
    file_frame = ttk.LabelFrame(main_frame, text="اختيار الملفات", padding="5")
    file_frame.pack(fill=tk.X, pady=5)
    
    ttk.Button(file_frame, text="اختيار ملفات", command=app.select_files).pack(side=tk.LEFT, padx=5)
    ttk.Button(file_frame, text="اختيار مجلد", command=app.select_folder).pack(side=tk.LEFT, padx=5)
    ttk.Label(file_frame, text="المجلد:").pack(side=tk.LEFT, padx=5)
    ttk.Entry(file_frame, textvariable=app.input_folder, width=50, state="readonly").pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
    
    # إطار عرض الملفات
    files_frame = ttk.LabelFrame(main_frame, text="الملفات المحددة", padding="5")
    files_frame.pack(fill=tk.BOTH, expand=True, pady=5)
    
    # إنشاء جدول لعرض الملفات
    files_tree_frame = ttk.Frame(files_frame)
    files_tree_frame.pack(fill=tk.BOTH, expand=True)
    
    # إنشاء شريط تمرير
    tree_scroll_y = ttk.Scrollbar(files_tree_frame, orient=tk.VERTICAL)
    tree_scroll_x = ttk.Scrollbar(files_tree_frame, orient=tk.HORIZONTAL)
    
    # إنشاء جدول العرض
    app.files_tree = ttk.Treeview(
        files_tree_frame, 
        columns=("file", "status", "columns", "selected_column", "sheet"),
        show="headings",
        yscrollcommand=tree_scroll_y.set,
        xscrollcommand=tree_scroll_x.set
    )
    
    # تكوين أعمدة الجدول
    app.files_tree.heading("file", text="اسم الملف")
    app.files_tree.heading("status", text="الحالة")
    app.files_tree.heading("columns", text="الأعمدة")
    app.files_tree.heading("selected_column", text="عامود المقارنة")
    app.files_tree.heading("sheet", text="ورقة العمل")
    
    app.files_tree.column("file", width=200)  # زيادة عرض عامود اسم الملف
    app.files_tree.column("status", width=80)
    app.files_tree.column("columns", width=250)
    app.files_tree.column("selected_column", width=120)
    app.files_tree.column("sheet", width=100)
    
    # ترتيب عناصر الجدول
    app.files_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    tree_scroll_y.config(command=app.files_tree.yview)
    tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
    tree_scroll_x.config(command=app.files_tree.xview)
    tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
    
    # إنشاء قائمة سياقية للجدول
    app.context_menu = tk.Menu(app.root, tearoff=0)
    app.context_menu.add_command(label="تحديد عامود للمقارنة", command=app.select_column_for_file)
    # تمت إزالة خيار "تحديد ورقة العمل" من القائمة السياقية
    app.context_menu.add_separator()
    app.context_menu.add_command(label="إزالة", command=app.remove_selected_file)
    
    # ربط النقر المزدوج بتحديد عامود المقارنة
    app.files_tree.bind("<Double-1>", app.select_column_for_file)
    
    # ربط النقر بالزر الأيمن بالقائمة السياقية
    app.files_tree.bind("<Button-3>", app.show_context_menu)
    
    # إطار اختيار العامود للمقارنة
    column_frame = ttk.LabelFrame(main_frame, text="خيارات المقارنة", padding="5")
    column_frame.pack(fill=tk.X, pady=5)
    
    ttk.Checkbutton(
        column_frame, 
        text="استخدام نفس العامود للمقارنة", 
        variable=app.use_same_column,
        command=app.toggle_column_selection
    ).pack(side=tk.LEFT, padx=5)
    
    # إطار العامود المشترك
    app.common_column_frame = ttk.Frame(column_frame)
    app.common_column_frame.pack(fill=tk.X, pady=5)
    
    ttk.Label(app.common_column_frame, text="العامود المشترك:").pack(side=tk.LEFT, padx=5)
    app.common_column_combobox = ttk.Combobox(
        app.common_column_frame, 
        textvariable=app.common_column,
        state="readonly",
        width=30
    )
    app.common_column_combobox.pack(side=tk.LEFT, padx=5)
    
    # إطار ملف الإخراج
    output_frame = ttk.LabelFrame(main_frame, text="ملف الإخراج", padding="5")
    output_frame.pack(fill=tk.X, pady=5)
    
    ttk.Label(output_frame, text="ملف النتائج:").pack(side=tk.LEFT, padx=5)
    ttk.Entry(output_frame, textvariable=app.output_file, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
    ttk.Button(output_frame, text="اختيار", command=app.select_output_file).pack(side=tk.LEFT, padx=5)
    
    # إطار التحكم
    control_frame = ttk.Frame(main_frame)
    control_frame.pack(fill=tk.X, pady=5)
    
    app.run_button = ttk.Button(control_frame, text="تشغيل المقارنة", command=app.run_comparison)
    app.run_button.pack(side=tk.LEFT, padx=5)
    
    ttk.Button(control_frame, text="مسح الكل", command=app.clear_all).pack(side=tk.LEFT, padx=5)
    
    # شريط التقدم
    app.progress = ttk.Progressbar(control_frame, mode="indeterminate", length=200)
    app.progress.pack(side=tk.LEFT, padx=10)
    
    # إطار سجل العمليات
    log_frame = ttk.LabelFrame(main_frame, text="سجل العمليات", padding="5")
    log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
    
    app.log_text = scrolledtext.ScrolledText(log_frame, height=10, wrap=tk.WORD)
    app.log_text.pack(fill=tk.BOTH, expand=True)
    app.log_text.configure(state="disabled")
    
    # إطار المعلومات
    info_frame = ttk.Frame(main_frame)
    info_frame.pack(fill=tk.X, pady=10)  # زيادة المساحة
    
    ttk.Button(info_frame, text="حول البرنامج", command=app.about).pack(side=tk.LEFT, padx=5)
    ttk.Button(info_frame, text="مساعدة", command=app.show_help).pack(side=tk.LEFT, padx=5)
    ttk.Button(info_frame, text="خروج", command=app.root.quit).pack(side=tk.LEFT, padx=5)
    
    # تعطيل بعض العناصر في البداية
    app.common_column_combobox.config(state="disabled")