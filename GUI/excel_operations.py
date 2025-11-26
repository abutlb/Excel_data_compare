#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

def read_excel_file(app, file_path, use_first_sheet=False):
    """قراءة ملف إكسل وعرض الأعمدة"""
    # قراءة أسماء أوراق العمل في الملف
    xl = pd.ExcelFile(file_path)
    sheet_names = xl.sheet_names
    app.file_sheets[file_path] = sheet_names
    
    # إذا كان هناك أكثر من ورقة ولم يتم تحديد استخدام الورقة الأولى
    if len(sheet_names) > 1 and not use_first_sheet:
        select_multiple_sheets(app, file_path=file_path)
    else:
        # استخدم الورقة الأولى
        selected_sheet = sheet_names[0]
        app.selected_sheets[file_path] = selected_sheet
        
        # قراءة البيانات من الورقة المحددة
        df = pd.read_excel(file_path, sheet_name=selected_sheet)
        columns = df.columns.tolist()
        app.file_columns[file_path] = columns
        
        # إضافة الملف إلى الجدول
        file_name = os.path.basename(file_path)
        app.files_tree.insert("", "end", values=(file_name, "تم القراءة", ", ".join(columns[:3]) + "...", "لم يتم التحديد", selected_sheet))
        
        app.log(f"تمت إضافة الملف: {file_name} (الورقة: {selected_sheet})")
        
        # تحديث قائمة الأعمدة المشتركة
        update_common_columns(app)

def update_common_columns(app):
    """تحديث قائمة الأعمدة المشتركة بين جميع الملفات"""
    if not app.file_columns:
        return
        
    # الحصول على الأعمدة المشتركة
    common_columns = set(next(iter(app.file_columns.values())))
    
    for columns in app.file_columns.values():
        common_columns &= set(columns)
    
    # تحديث قائمة الأعمدة المشتركة
    app.common_column_combobox['values'] = list(common_columns)
    
    if common_columns:
        app.common_column.set(next(iter(common_columns)))
        app.common_column_combobox.config(state="readonly")
        app.log(f"تم العثور على {len(common_columns)} عامود مشترك بين جميع الملفات.")
    else:
        app.common_column.set("")
        app.common_column_combobox.config(state="disabled")
        app.log("لا توجد أعمدة مشتركة بين جميع الملفات.")

def select_multiple_sheets(app, event=None, file_path=None):
    """اختيار ورقة عمل أو أكثر لملف محدد"""
    # إذا لم يتم تمرير مسار الملف، استخدم العنصر المحدد
    if file_path is None:
        selection = app.files_tree.selection()
        if not selection:
            return
            
        item = app.files_tree.item(selection[0])
        file_name = item['values'][0]
        
        # البحث عن مسار الملف
        file_path = None
        for path in app.excel_files:
            if os.path.basename(path) == file_name:
                file_path = path
                break
                
        if not file_path:
            return
    
    # التحقق من وجود أوراق عمل للملف
    if file_path not in app.file_sheets:
        try:
            xl = pd.ExcelFile(file_path)
            app.file_sheets[file_path] = xl.sheet_names
        except Exception as e:
            messagebox.showerror("خطأ", f"تعذر قراءة أوراق العمل: {str(e)}")
            return
    
    sheet_names = app.file_sheets[file_path]
    if not sheet_names:
        messagebox.showinfo("تنبيه", "لا توجد أوراق عمل في هذا الملف.")
        return
    
    # إنشاء نافذة منبثقة لاختيار ورقة العمل
    popup = tk.Toplevel(app.root)
    popup.title(f"اختر ورقة العمل - {os.path.basename(file_path)}")
    popup.geometry("500x500")  # زيادة حجم النافذة
    popup.transient(app.root)
    popup.grab_set()
    
    # إنشاء قائمة لأوراق العمل
    ttk.Label(popup, text="اختر ورقة العمل أو أكثر:").pack(pady=10)
    
    # إطار للقائمة وشريط التمرير
    list_frame = ttk.Frame(popup)
    list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
    
    # استخدام اختيار متعدد
    sheets_listbox = tk.Listbox(list_frame, width=40, height=15, selectmode=tk.MULTIPLE)
    sheets_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    
    # إضافة شريط تمرير
    scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=sheets_listbox.yview)
    sheets_listbox.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    # إضافة أوراق العمل إلى القائمة
    for sheet in sheet_names:
        sheets_listbox.insert(tk.END, sheet)
    
    # إطار للأزرار
    buttons_frame = ttk.Frame(popup)
    buttons_frame.pack(fill=tk.X, pady=20)  # زيادة المساحة
    
    def on_select():
        """عند اختيار ورقة"""
        selections = sheets_listbox.curselection()
        if not selections:
            messagebox.showinfo("تنبيه", "يرجى اختيار ورقة عمل واحدة على الأقل.")
            return
        
        # مسح الملف الحالي من القائمة إذا كان موجوداً
        original_file_name = os.path.basename(file_path)
        for item in app.files_tree.get_children():
            if app.files_tree.item(item)['values'][0] == original_file_name:
                app.files_tree.delete(item)
                break
        
        # إضافة كل ورقة مختارة كملف منفصل
        for idx in selections:
            sheet = sheet_names[idx]
            try:
                # قراءة البيانات من الورقة المحددة
                df = pd.read_excel(file_path, sheet_name=sheet)
                columns = df.columns.tolist()
                
                # إنشاء اسم فريد للملف مع ورقة العمل
                display_name = f"{original_file_name} - {sheet}"
                
                # تخزين البيانات
                virtual_path = f"{file_path}::{sheet}"  # استخدام مسار افتراضي لتمييز الورقة
                app.excel_files.append(virtual_path)
                app.file_columns[virtual_path] = columns
                app.selected_sheets[virtual_path] = sheet
                
                # إضافة إلى الجدول
                app.files_tree.insert("", "end", values=(display_name, "تم القراءة", ", ".join(columns[:3]) + "...", "لم يتم التحديد", sheet))
                
                app.log(f"تمت إضافة الملف: {display_name}")
            except Exception as e:
                app.log(f"خطأ في قراءة الورقة '{sheet}': {str(e)}")
        
        # تحديث قائمة الأعمدة المشتركة
        update_common_columns(app)
        popup.destroy()
    
    def on_cancel():
        """إلغاء الاختيار"""
        popup.destroy()
    
    ttk.Button(buttons_frame, text="اختيار", command=on_select).pack(side=tk.LEFT, padx=20)
    ttk.Button(buttons_frame, text="إلغاء", command=on_cancel).pack(side=tk.LEFT)
    
    # وضع النافذة في المنتصف
    popup.update_idletasks()
    width = popup.winfo_width()
    height = popup.winfo_height()
    x = (popup.winfo_screenwidth() // 2) - (width // 2)
    y = (popup.winfo_screenheight() // 2) - (height // 2)
    popup.geometry('{}x{}+{}+{}'.format(width, height, x, y))

def select_sheet_for_file(app, event=None, file_path=None):
    """اختيار ورقة عمل لملف محدد - تحويل إلى اختيار العامود"""
    # بدلاً من اختيار الورقة، نقوم بتشغيل اختيار العامود مباشرة
    select_column_for_file(app, event)

def select_column_for_file(app, event=None):
    """اختيار عامود للمقارنة لملف محدد"""
    if app.use_same_column.get():
        messagebox.showinfo("تنبيه", "يرجى إلغاء تحديد خيار 'استخدام نفس العامود للمقارنة' أولاً.")
        return
        
    # الحصول على العنصر المحدد
    selection = app.files_tree.selection()
    if not selection:
        return
        
    item = app.files_tree.item(selection[0])
    file_name = item['values'][0]
    
    # البحث عن مسار الملف
    file_path = None
    for path in app.excel_files:
        if os.path.basename(path) == file_name or path.endswith(f"::{file_name.split(' - ')[-1]}"):
            file_path = path
            break
            
    if not file_path or file_path not in app.file_columns:
        return
        
    # إنشاء نافذة منبثقة لاختيار العامود
    popup = tk.Toplevel(app.root)
    popup.title(f"اختر عامود للمقارنة - {file_name}")
    popup.geometry("500x400")  # زيادة حجم النافذة
    popup.transient(app.root)
    popup.grab_set()
    
    # إنشاء قائمة للأعمدة
    ttk.Label(popup, text="اختر العامود للمقارنة:").pack(pady=10)
    
    # إطار للقائمة وشريط التمرير
    list_frame = ttk.Frame(popup)
    list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
    
    columns_listbox = tk.Listbox(list_frame, width=40, height=15)
    columns_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    
    # إضافة شريط تمرير
    scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=columns_listbox.yview)
    columns_listbox.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    # إضافة الأعمدة إلى القائمة
    for column in app.file_columns[file_path]:
        columns_listbox.insert(tk.END, column)
        
    # تحديد العامود الحالي إذا كان موجوداً
    if file_path in app.selected_columns:
        try:
            index = app.file_columns[file_path].index(app.selected_columns[file_path])
            columns_listbox.selection_set(index)
            columns_listbox.see(index)
        except ValueError:
            pass
    
    # إطار للأزرار
    buttons_frame = ttk.Frame(popup)
    buttons_frame.pack(fill=tk.X, pady=20)  # زيادة المساحة
    
    def on_select():
        """عند اختيار عامود"""
        selection = columns_listbox.curselection()
        if selection:
            column = app.file_columns[file_path][selection[0]]
            app.selected_columns[file_path] = column
            app.update_file_display()
            app.log(f"تم تحديد العامود '{column}' للملف {file_name}")
            popup.destroy()
    
    def on_cancel():
        """إلغاء الاختيار"""
        popup.destroy()
    
    ttk.Button(buttons_frame, text="اختيار", command=on_select).pack(side=tk.LEFT, padx=20)
    ttk.Button(buttons_frame, text="إلغاء", command=on_cancel).pack(side=tk.LEFT)
    
    # وضع النافذة في المنتصف
    popup.update_idletasks()
    width = popup.winfo_width()
    height = popup.winfo_height()
    x = (popup.winfo_screenwidth() // 2) - (width // 2)
    y = (popup.winfo_screenheight() // 2) - (height // 2)
    popup.geometry('{}x{}+{}+{}'.format(width, height, x, y))