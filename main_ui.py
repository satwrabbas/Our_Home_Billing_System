import customtkinter as ctk
from tkinter import messagebox
from datetime import datetime
from excel_manager import ExcelManager

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("نظام إدارة مبيعات الشقق والأقساط - النسخة الاحترافية")
        self.geometry("600x550")
        self.resizable(False, False)

        self.manager = ExcelManager()

        self.tabview = ctk.CTkTabview(self, width=550, height=500)
        self.tabview.pack(padx=20, pady=20)

        self.tab_payment = self.tabview.add("تسجيل دفعة (قسط)")
        self.tab_allocate = self.tabview.add("تخصيص شقة لعميل")

        self.setup_payment_tab()
        self.setup_allocate_tab()
        
        # استدعاء دالة تحديث القوائم عند فتح التطبيق
        self.refresh_dropdowns()

    def setup_payment_tab(self):
        ctk.CTkLabel(self.tab_payment, text=":اختر العميل", font=("Arial", 14, "bold")).pack(pady=(15, 5))
        # استخدام قائمة منسدلة بدلاً من إدخال النص
        self.combo_payment_client = ctk.CTkOptionMenu(self.tab_payment, width=300, font=("Arial", 14), justify="right")
        self.combo_payment_client.pack(pady=5)

        ctk.CTkLabel(self.tab_payment, text=":المبلغ (ليرة سورية)", font=("Arial", 14, "bold")).pack(pady=(10, 5))
        self.amount_entry = ctk.CTkEntry(self.tab_payment, width=300, justify="center")
        self.amount_entry.pack(pady=5)

        ctk.CTkLabel(self.tab_payment, text=":تاريخ الدفعة", font=("Arial", 14, "bold")).pack(pady=(10, 5))
        self.date_entry = ctk.CTkEntry(self.tab_payment, width=300, justify="center")
        self.date_entry.insert(0, datetime.now().strftime("%Y/%m/%d"))
        self.date_entry.pack(pady=5)

        ctk.CTkLabel(self.tab_payment, text=":ملاحظات الدفعة (اختياري)", font=("Arial", 14, "bold")).pack(pady=(10, 5))
        self.notes_entry = ctk.CTkEntry(self.tab_payment, width=300, justify="right")
        self.notes_entry.pack(pady=5)

        self.btn_save_payment = ctk.CTkButton(self.tab_payment, text="حفظ وإضافة الدفعة", font=("Arial", 15, "bold"), fg_color="green", hover_color="darkgreen", command=self.process_payment)
        self.btn_save_payment.pack(pady=30)

    def setup_allocate_tab(self):
        ctk.CTkLabel(self.tab_allocate, text=":اختر العميل (غير المتخصصين فقط)", font=("Arial", 14, "bold")).pack(pady=(15, 5))
        # قائمة منسدلة خاصة بالعملاء الذين لم يتخصصوا بعد
        self.combo_alloc_client = ctk.CTkOptionMenu(self.tab_allocate, width=300, font=("Arial", 14), justify="right")
        self.combo_alloc_client.pack(pady=5)

        ctk.CTkLabel(self.tab_allocate, text=":مساحة الشقة (بالمتر المربع)", font=("Arial", 14, "bold")).pack(pady=(10, 5))
        self.area_entry = ctk.CTkEntry(self.tab_allocate, width=300, justify="center")
        self.area_entry.pack(pady=5)

        ctk.CTkLabel(self.tab_allocate, text=":رقم أو معامل الطابق", font=("Arial", 14, "bold")).pack(pady=(10, 5))
        self.floor_entry = ctk.CTkEntry(self.tab_allocate, width=300, justify="center")
        self.floor_entry.pack(pady=5)

        self.btn_allocate = ctk.CTkButton(self.tab_allocate, text="تخصيص الشقة ونقل الملف", font=("Arial", 15, "bold"), command=self.process_allocation)
        self.btn_allocate.pack(pady=40)

    # ------------------ دالة تحديث القوائم ------------------ #
    def refresh_dropdowns(self):
        """تقرأ الملفات من المجلدات وتحدث القوائم المنسدلة"""
        all_clients = self.manager.get_all_clients()
        unallocated_clients = self.manager.get_unallocated_clients()

        # تحديث قائمة الدفع (تظهر جميع العملاء)
        if all_clients:
            self.combo_payment_client.configure(values=all_clients)
            self.combo_payment_client.set(all_clients[0])
        else:
            self.combo_payment_client.configure(values=["لا يوجد عملاء"])
            self.combo_payment_client.set("لا يوجد عملاء")

        # تحديث قائمة التخصيص (تظهر فقط غير المتخصصين)
        if unallocated_clients:
            self.combo_alloc_client.configure(values=unallocated_clients)
            self.combo_alloc_client.set(unallocated_clients[0])
        else:
            self.combo_alloc_client.configure(values=["لا يوجد عملاء"])
            self.combo_alloc_client.set("لا يوجد عملاء")

    # ------------------ الوظائف التشغيلية ------------------ #
    def process_payment(self):
        client_name = self.combo_payment_client.get()
        if client_name == "لا يوجد عملاء":
            messagebox.showwarning("تنبيه", "لا يوجد ملفات عملاء مسجلة!")
            return

        amount_str = self.amount_entry.get().strip()
        date = self.date_entry.get().strip()
        notes = self.notes_entry.get().strip()

        if not amount_str:
            messagebox.showwarning("تنبيه", "يرجى إدخال المبلغ!")
            return

        try:
            amount = float(amount_str)
        except ValueError:
            messagebox.showerror("خطأ", "يجب أن يكون المبلغ رقماً صحيحاً!")
            return

        success, message = self.manager.add_payment(client_name, amount, date, notes)
        
        if success:
            messagebox.showinfo("نجاح", message)
            self.amount_entry.delete(0, 'end')
            self.notes_entry.delete(0, 'end')
        else:
            messagebox.showerror("خطأ", message)

    def process_allocation(self):
        client_name = self.combo_alloc_client.get()
        if client_name == "لا يوجد عملاء":
            messagebox.showwarning("تنبيه", "لا يوجد عملاء بانتظار التخصيص!")
            return

        area_str = self.area_entry.get().strip()
        floor_str = self.floor_entry.get