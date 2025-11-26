#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import threading
from datetime import datetime

from ui_components import create_ui
from excel_operations import read_excel_file, update_common_columns, select_multiple_sheets
from excel_operations import select_sheet_for_file as ops_select_sheet, select_column_for_file as ops_select_column
from report_generator import perform_comparison

class ExcelComparerApp:
    """تطبيق مقارنة ملفات الإكسل"""
    
    def __init__(self, root):
        """تهيئة التطبيق"""
        self.root = root
        self.root.geometry("1000x700")  # زيادة حجم النافذة الرئيسية
        
        # متغيرات لتخزين البيانات
        self.excel_files = []  # قائمة بمسارات الملفات
        self.file_columns = {}  # قاموس للأعمدة في كل ملف
        self.selected_columns = {}  # قاموس للأعمدة المختارة للمقارنة
        self.file_sheets = {}  # قاموس لأوراق العمل في كل ملف
        self.selected_sheets = {}  # قاموس لأوراق العمل المختارة
        self.is_running = False  # حالة تشغيل المقارنة
        self.main_file = None  # الملف الأساسي للمقارنة
        
        # متغيرات Tkinter
        self.input_folder = tk.StringVar()
        self.output_file = tk.StringVar()
        self.common_column = tk.StringVar()
        self.use_same_column = tk.BooleanVar(value=True)
        
        # إنشاء واجهة المستخدم
        create_ui(self)
        
        # طباعة رسالة ترحيبية
        self.log("مرحباً بك في برنامج مقارنة ملفات الإكسل!")
        self.log("يرجى اختيار الملفات أو المجلد للبدء.")
    
    def log(self, message):
        """إضافة رسالة إلى سجل العمليات"""
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")
        # تحديث الواجهة
        self.root.update_idletasks()
    
    def select_files(self):
        """اختيار ملفات محددة"""
        files = filedialog.askopenfilenames(
            title="اختر ملفات الإكسل",
            filetypes=[("ملفات إكسل", "*.xlsx *.xls")]
        )
        
        if not files:
            return
            
        # إضافة الملفات المختارة
        for file_path in files:
            if file_path not in self.excel_files:
                self.excel_files.append(file_path)
                self.input_folder.set(os.path.dirname(file_path))
                
                # محاولة قراءة الملف وعرض الأعمدة
                try:
                    read_excel_file(self, file_path)
                except Exception as e:
                    self.files_tree.insert("", "end", values=(os.path.basename(file_path), "خطأ", str(e), "-", "-"))
                    self.log(f"خطأ في قراءة الملف {os.path.basename(file_path)}: {str(e)}")
    
    def select_folder(self):
        """اختيار مجلد يحتوي على ملفات إكسل"""
        folder = filedialog.askdirectory(title="اختر المجلد الذي يحتوي على ملفات الإكسل")
        
        if not folder:
            return
            
        self.input_folder.set(folder)
        
        # البحث عن ملفات إكسل في المجلد
        excel_files = [os.path.join(folder, f) for f in os.listdir(folder) 
                      if f.endswith(('.xlsx', '.xls'))]
        
        if not excel_files:
            messagebox.showinfo("تنبيه", "لم يتم العثور على ملفات إكسل في المجلد المحدد.")
            return
            
        # مسح الملفات الحالية
        self.excel_files = []
        self.file_columns = {}
        self.selected_columns = {}
        self.file_sheets = {}
        self.selected_sheets = {}
        self.files_tree.delete(*self.files_tree.get_children())
        self.main_file = None  # إعادة تعيين الملف الأساسي
        
        # إضافة الملفات الجديدة
        for file_path in excel_files:
            self.excel_files.append(file_path)
            
            # محاولة قراءة الملف وعرض الأعمدة
            try:
                read_excel_file(self, file_path, use_first_sheet=True)
            except Exception as e:
                self.files_tree.insert("", "end", values=(os.path.basename(file_path), "خطأ", str(e), "-", "-"))
                self.log(f"خطأ في قراءة الملف {os.path.basename(file_path)}: {str(e)}")
        
        # تحديث قائمة الأعمدة المشتركة
        update_common_columns(self)
    
    # إضافة الدوال المفقودة
    def select_sheet_for_file(self, event=None, file_path=None):
        """اختيار ورقة عمل لملف محدد"""
        ops_select_sheet(self, event, file_path)
    
    def select_column_for_file(self, event=None):
        """اختيار عامود للمقارنة لملف محدد"""
        ops_select_column(self, event)
    
    def set_as_main_file(self):
        """تعيين الملف المحدد كملف أساسي للمقارنة"""
        selection = self.files_tree.selection()
        if not selection:
            return
            
        item = self.files_tree.item(selection[0])
        file_name = item['values'][0]
        
        # البحث عن مسار الملف
        file_path = None
        for path in self.excel_files:
            # التعامل مع الملفات العادية والملفات ذات المسار الافتراضي (للأوراق المتعددة)
            if os.path.basename(path) == file_name or ("::" in path and file_name.endswith(path.split("::")[-1])):
                file_path = path
                break
                
        if not file_path:
            return
        
        # تحديث الملف الأساسي
        self.main_file = file_path
        
        # تحديث العرض في الجدول
        for item in self.files_tree.get_children():
            current_values = self.files_tree.item(item)['values']
            current_file_name = current_values[0]
            
            # تحديث حالة الملف (إضافة علامة للملف الأساسي)
            if current_file_name == file_name:
                self.files_tree.item(item, values=(current_values[0], "ملف أساسي", current_values[2], current_values[3], current_values[4]))
            elif current_values[1] == "ملف أساسي":
                self.files_tree.item(item, values=(current_values[0], "تم القراءة", current_values[2], current_values[3], current_values[4]))
        
        self.log(f"تم تعيين الملف '{file_name}' كملف أساسي للمقارنة")
    
    def toggle_column_selection(self):
        """تبديل وضع اختيار العامود"""
        if self.use_same_column.get():
            self.common_column_frame.pack(fill=tk.X, pady=5)
            # تحديث الأعمدة المختارة لجميع الملفات
            if self.common_column.get():
                for file_path in self.excel_files:
                    if file_path in self.file_columns and self.common_column.get() in self.file_columns[file_path]:
                        self.selected_columns[file_path] = self.common_column.get()
                
                # تحديث العرض
                self.update_file_display()
        else:
            self.common_column_frame.pack_forget()
    
    def update_file_display(self):
        """تحديث عرض الملفات في الجدول"""
        for item in self.files_tree.get_children():
            file_name = self.files_tree.item(item)['values'][0]
            
            # البحث عن مسار الملف
            file_path = None
            for path in self.excel_files:
                # التعامل مع الملفات العادية والملفات ذات المسار الافتراضي (للأوراق المتعددة)
                if os.path.basename(path) == file_name or ("::" in path and file_name.endswith(path.split("::")[-1])):
                    file_path = path
                    break
            
            if file_path:
                selected_column = "لم يتم التحديد"
                if file_path in self.selected_columns:
                    selected_column = self.selected_columns[file_path]
                elif self.use_same_column.get() and self.common_column.get():
                    if file_path in self.file_columns and self.common_column.get() in self.file_columns[file_path]:
                        selected_column = self.common_column.get()
                        self.selected_columns[file_path] = selected_column
                
                selected_sheet = "الافتراضية"
                if file_path in self.selected_sheets:
                    selected_sheet = self.selected_sheets[file_path]
                elif "::" in file_path:
                    selected_sheet = file_path.split("::")[1]
                
                # تحديث حالة الملف (الحفاظ على علامة الملف الأساسي)
                status = self.files_tree.item(item)['values'][1]
                if file_path == self.main_file:
                    status = "ملف أساسي"
                elif status != "خطأ" and status != "ملف أساسي":
                    status = "تم القراءة"
                
                # تحديث العرض
                columns_text = self.files_tree.item(item)['values'][2]
                self.files_tree.item(item, values=(file_name, status, columns_text, selected_column, selected_sheet))
    
    def select_output_file(self):
        """اختيار ملف الإخراج"""
        file_path = filedialog.asksaveasfilename(
            title="حفظ ملف المقارنة",
            defaultextension=".xlsx",
            filetypes=[("ملف إكسل", "*.xlsx")]
        )
        
        if file_path:
            self.output_file.set(file_path)
    
    def clear_all(self):
        """مسح جميع البيانات"""
        if messagebox.askyesno("تأكيد", "هل أنت متأكد من مسح جميع البيانات؟"):
            self.excel_files = []
            self.file_columns = {}
            self.selected_columns = {}
            self.file_sheets = {}
            self.selected_sheets = {}
            self.main_file = None  # إعادة تعيين الملف الأساسي
            self.files_tree.delete(*self.files_tree.get_children())
            self.input_folder.set("")
            self.common_column.set("")
            self.common_column_combobox.config(state="disabled")
            self.log_text.configure(state="normal")
            self.log_text.delete(1.0, tk.END)
            self.log_text.configure(state="disabled")
            self.log("تم مسح جميع البيانات.")
    
    def run_comparison(self):
        """تشغيل عملية المقارنة"""
        if self.is_running:
            return
            
        # التحقق من وجود ملفات في القائمة
        if not self.files_tree.get_children():
            messagebox.showwarning("تحذير", "لم يتم اختيار أي ملفات للمقارنة.")
            return
        
        # جمع معلومات الملفات من القائمة المعروضة
        files_info = []
        for item in self.files_tree.get_children():
            values = self.files_tree.item(item)['values']
            file_name = values[0]
            status = values[1]
            selected_column = values[3]
            
            # تخطي الملفات التي بها خطأ
            if status == "خطأ":
                continue
                
            # البحث عن مسار الملف
            file_path = None
            for path in self.excel_files:
                if os.path.basename(path) == file_name or ("::" in path and file_name.endswith(path.split("::")[-1])):
                    file_path = path
                    break
            
            if file_path:
                files_info.append({
                    'name': file_name,
                    'path': file_path,
                    'column': selected_column,
                    'is_main': (status == "ملف أساسي")
                })
        
        if not files_info:
            messagebox.showwarning("تحذير", "لم يتم العثور على ملفات صالحة للمقارنة.")
            return
            
        # التحقق من اختيار العامود للمقارنة
        if self.use_same_column.get():
            if not self.common_column.get():
                messagebox.showwarning("تحذير", "يرجى اختيار العامود المشترك للمقارنة.")
                return
            
            # تعيين العامود المشترك لجميع الملفات
            comparison_column = self.common_column.get()
        else:
            # التحقق من أن جميع الملفات لها عامود محدد
            missing_columns = []
            for file_info in files_info:
                if file_info['column'] == "لم يتم التحديد":
                    missing_columns.append(file_info['name'])
            
            if missing_columns:
                messagebox.showwarning("تحذير", f"يرجى تحديد عامود للمقارنة للملفات التالية:\n{', '.join(missing_columns)}")
                return
                
            # إنشاء قاموس للأعمدة المختارة
            comparison_column = {}
            for file_info in files_info:
                comparison_column[file_info['path']] = file_info['column']
        
        # التحقق من اسم ملف الإخراج
        output_file = self.output_file.get()
        if not output_file:
            messagebox.showwarning("تحذير", "يرجى تحديد اسم ملف الإخراج.")
            return
            
        if not output_file.endswith('.xlsx'):
            output_file += '.xlsx'
            self.output_file.set(output_file)
        
        # تعيين الملف الأساسي إذا تم تحديده
        main_file_info = None
        for file_info in files_info:
            if file_info['is_main']:
                main_file_info = file_info
                break
        
        # تعطيل الواجهة أثناء المعالجة
        self.is_running = True
        self.run_button.config(state="disabled")
        self.progress.start()
        
        # تشغيل المقارنة في خيط منفصل
        thread = threading.Thread(
            target=perform_comparison, 
            args=(self, comparison_column, output_file, main_file_info)
        )
        thread.daemon = True
        thread.start()
    
    def finish_comparison(self, success=True):
        """إنهاء عملية المقارنة وإعادة تفعيل الواجهة"""
        self.is_running = False
        self.run_button.config(state="normal")
        self.progress.stop()
        
        if success:
            messagebox.showinfo("اكتمال", f"تمت عملية المقارنة بنجاح وتم حفظ النتائج في:\n{self.output_file.get()}")
        else:
            messagebox.showerror("خطأ", "حدث خطأ أثناء عملية المقارنة. يرجى التحقق من سجل العمليات.")
    
    def remove_selected_file(self):
        """إزالة الملف المحدد من القائمة"""
        selection = self.files_tree.selection()
        if not selection:
            return
            
        item = self.files_tree.item(selection[0])
        file_name = item['values'][0]
        
        # البحث عن مسار الملف وإزالته
        for path in self.excel_files[:]:
            # التعامل مع الملفات العادية والملفات ذات المسار الافتراضي (للأوراق المتعددة)
            if os.path.basename(path) == file_name or ("::" in path and file_name.endswith(path.split("::")[-1])):
                self.excel_files.remove(path)
                if path in self.file_columns:
                    del self.file_columns[path]
                if path in self.selected_columns:
                    del self.selected_columns[path]
                if path in self.file_sheets:
                    del self.file_sheets[path]
                if path in self.selected_sheets:
                    del self.selected_sheets[path]
                
                # إذا كان هذا هو الملف الأساسي، قم بإعادة تعيينه
                if path == self.main_file:
                    self.main_file = None
                
                break
                
        # إزالة العنصر من الجدول
        self.files_tree.delete(selection[0])
        self.log(f"تمت إزالة الملف: {file_name}")
        
        # تحديث قائمة الأعمدة المشتركة
        update_common_columns(self)
    
    def show_context_menu(self, event):
        """عرض القائمة السياقية عند النقر بالزر الأيمن"""
        selection = self.files_tree.identify_row(event.y)
        if selection:
            self.files_tree.selection_set(selection)
            self.context_menu.post(event.x_root, event.y_root)
    
    def about(self):
        """عرض معلومات حول البرنامج"""
        about_window = tk.Toplevel(self.root)
        about_window.title("حول البرنامج")
        about_window.geometry("500x350")  # زيادة حجم النافذة
        about_window.resizable(False, False)
        about_window.transient(self.root)
        about_window.grab_set()
        
        # إضافة المعلومات
        ttk.Label(about_window, text="برنامج مقارنة ملفات الإكسل", font=("Arial", 16, "bold")).pack(pady=20)
        ttk.Label(about_window, text="الإصدار 1.2").pack()
        ttk.Label(about_window, text="").pack()
        ttk.Label(about_window, text="برنامج مفتوح المصدر لمقارنة ملفات الإكسل وتحديد السجلات الفريدة").pack(pady=5)
        ttk.Label(about_window, text="يمكنك مقارنة ملفين أو أكثر وتحديد السجلات المشتركة والفريدة").pack(pady=5)
        ttk.Label(about_window, text="يدعم اختيار أوراق عمل متعددة من نفس الملف").pack(pady=5)
        ttk.Label(about_window, text="يمكن تحديد ملف أساسي لاستخدام مسميات أعمدته كمرجع").pack(pady=5)
        
        # زر الإغلاق
        ttk.Button(about_window, text="إغلاق", command=about_window.destroy).pack(pady=30)
        
        # وضع النافذة في المنتصف
        about_window.update_idletasks()
        width = about_window.winfo_width()
        height = about_window.winfo_height()
        x = (about_window.winfo_screenwidth() // 2) - (width // 2)
        y = (about_window.winfo_screenheight() // 2) - (height // 2)
        about_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def show_help(self):
        """عرض مساعدة البرنامج"""
        help_window = tk.Toplevel(self.root)
        help_window.title("مساعدة")
        help_window.geometry("700x600")  # زيادة حجم النافذة
        help_window.transient(self.root)
        help_window.grab_set()
        
        # إنشاء نص المساعدة مع شريط تمرير
        help_text = scrolledtext.ScrolledText(help_window, wrap=tk.WORD, width=80, height=30)
        help_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        
        # إضافة محتوى المساعدة
        help_content = """
        كيفية استخدام برنامج مقارنة ملفات الإكسل:
        
        1. اختيار الملفات:
           - انقر على زر "اختيار ملفات" لاختيار ملفات إكسل محددة.
           - أو انقر على زر "اختيار مجلد" لاختيار مجلد يحتوي على ملفات إكسل.
           - عند اختيار ملف يحتوي على أكثر من ورقة عمل، ستظهر نافذة تتيح لك اختيار ورقة واحدة أو أكثر.
           - كل ورقة عمل تختارها ستظهر كملف منفصل في القائمة.
        
        2. تحديد الملف الأساسي (اختياري):
           - يمكنك تحديد ملف كملف أساسي للمقارنة عن طريق النقر بزر الفأرة الأيمن واختيار "تعيين كملف أساسي".
           - سيتم استخدام مسميات الأعمدة في الملف الأساسي كمرجع للمقارنة.
        
        3. اختيار العامود للمقارنة:
           - يمكنك استخدام نفس العامود للمقارنة في جميع الملفات عن طريق تحديد "استخدام نفس العامود للمقارنة".
           - أو يمكنك اختيار عامود مختلف لكل ملف عن طريق إلغاء تحديد الخيار ثم النقر المزدوج على الملف أو النقر بزر الفأرة الأيمن واختيار "تحديد عامود للمقارنة".
        
        4. تحديد ملف الإخراج:
           - انقر على زر "اختيار" بجانب "ملف الإخراج" لتحديد اسم ومكان ملف النتائج.
        
        5. تشغيل المقارنة:
           - بعد إعداد كل الخيارات، انقر على زر "تشغيل المقارنة" لبدء العملية.
           - سيتم عرض تقدم العملية في سجل العمليات أسفل النافذة.
        
        6. النتائج:
           - سيتم إنشاء ملف إكسل يحتوي على عدة أوراق عمل:
             * ملخص المقارنة: يعرض ملخصًا للعملية والإحصائيات.
             * جميع السجلات الفريدة: يعرض جميع السجلات الفريدة من جميع الملفات.
             * ورقة لكل ملف: تعرض السجلات الفريدة الموجودة في هذا الملف فقط.
        
        ملاحظات:
        - يمكنك إزالة ملف من القائمة بالنقر عليه بزر الفأرة الأيمن واختيار "إزالة".
        - يمكنك مسح جميع البيانات بالنقر على زر "مسح الكل".
        - يتم عرض سجل العمليات في الأسفل لمتابعة تقدم العملية.
        - النقر المزدوج على أي ملف سيفتح نافذة اختيار عامود المقارنة.
        - عند تحديد ملف أساسي، سيظهر بحالة "ملف أساسي" في القائمة.
        """
        
        help_text.insert(tk.END, help_content)
        help_text.configure(state="disabled")
        
        # زر الإغلاق
        ttk.Button(help_window, text="إغلاق", command=help_window.destroy).pack(pady=20)
        
        # وضع النافذة في المنتصف
        help_window.update_idletasks()
        width = help_window.winfo_width()
        height = help_window.winfo_height()
        x = (help_window.winfo_screenwidth() // 2) - (width // 2)
        y = (help_window.winfo_screenheight() // 2) - (height // 2)
        help_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))