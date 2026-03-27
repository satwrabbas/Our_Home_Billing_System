# gui_app.py
import customtkinter as ctk
from tkinter import messagebox
from datetime import datetime
from file_manager import FileManager
from excel_handler import ExcelHandler

# استدعاء مكتبات إصلاح اللغة العربية
import arabic_reshaper
from bidi.algorithm import get_display

def ar(text):
    """دالة مساعدة لربط وعكس الحروف العربية لتظهر بشكل صحيح في الواجهة"""
    if not text:
        return ""
    reshaped_text = arabic_reshaper.reshape(str(text))
    bidi_text = get_display(reshaped_text)
    return bidi_text

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class RealEstateApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # استخدام دالة ar() لعناوين النوافذ والأزرار
        self.title(ar("نظام إدارة العقارات والمبيعات الذكي"))
        self.geometry("850x600")
        
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # ================= القائمة الجانبية =================
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(5, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text=ar("بيتنا العقارية") + "\nOur Home", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 30))

        self.btn_add_payment = ctk.CTkButton(self.sidebar_frame, text=ar("إضافة دفعة مالية"), command=self.show_payment_frame)
        self.btn_add_payment.grid(row=1, column=0, padx=20, pady=10)

        self.btn_allocate = ctk.CTkButton(self.sidebar_frame, text=ar("تخصيص شقة"), command=self.show_allocate_frame)
        self.btn_allocate.grid(row=2, column=0, padx=20, pady=10)

        self.btn_receipt = ctk.CTkButton(self.sidebar_frame, text=ar("استخراج إيصال (PDF)"), fg_color="#d9534f", hover_color="#c9302c", command=self.show_receipt_frame)
        self.btn_receipt.grid(row=3, column=0, padx=20, pady=10)

        # ================= الشاشة الرئيسية =================
        self.main_frame = ctk.CTkFrame(self, corner_radius=15)
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)

        self.show_payment_frame()

    def clear_main_frame(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

    # ------------------ 1. شاشة الدفعات ------------------
    def show_payment_frame(self):
        self.clear_main_frame()
        title = ctk.CTkLabel(self.main_frame, text=ar("تسجيل دفعة مالية جديدة"), font=ctk.CTkFont(size=24, weight="bold"))
        title.pack(pady=(30, 20))

        # القاموس الذكي لحل مشكلة اللغة العربية في القائمة المنسدلة
        clients = FileManager.get_all_clients()
        if clients:
            self.client_map = {ar(c): c for c in clients}
        else:
            self.client_map = {ar("لا يوجد عملاء"): ""}

        self.combo_client = ctk.CTkComboBox(self.main_frame, values=list(self.client_map.keys()), justify="center", width=400, height=40)
        self.combo_client.set(ar("...اختر العميل..."))
        self.combo_client.pack(pady=10)

        self.entry_note = ctk.CTkEntry(self.main_frame, placeholder_text=ar("البيان (مثال: القسط 34)"), justify="center", width=400, height=40)
        self.entry_note.pack(pady=10)

        self.entry_date = ctk.CTkEntry(self.main_frame, justify="center", width=400, height=40)
        self.entry_date.insert(0, datetime.now().strftime("%Y/%m/%d"))
        self.entry_date.pack(pady=10)

        self.entry_syp = ctk.CTkEntry(self.main_frame, placeholder_text=ar("المبلغ بالليرة (أرقام فقط)"), justify="center", width=400, height=40)
        self.entry_syp.pack(pady=10)

        self.entry_usd = ctk.CTkEntry(self.main_frame, placeholder_text=ar("المبلغ بالدولار (اختياري)"), justify="center", width=400, height=40)
        self.entry_usd.pack(pady=10)

        self.btn_save = ctk.CTkButton(self.main_frame, text=ar("حفظ وترحيل"), font=ctk.CTkFont(size=16, weight="bold"), height=50, width=400, fg_color="#28a745", hover_color="#218838", command=self.process_payment)
        self.btn_save.pack(pady=(30, 10))

    # ------------------ 2. شاشة التخصيص ------------------
    def show_allocate_frame(self):
        self.clear_main_frame()
        title = ctk.CTkLabel(self.main_frame, text=ar("تخصيص شقة"), font=ctk.CTkFont(size=24, weight="bold"))
        title.pack(pady=(30, 20))
        
        clients = FileManager.get_unallocated_clients()
        if clients:
            self.allocate_map = {ar(c): c for c in clients}
        else:
            self.allocate_map = {ar("لا يوجد عملاء"): ""}

        self.combo_allocate = ctk.CTkComboBox(self.main_frame, values=list(self.allocate_map.keys()), justify="center", width=400, height=40)
        self.combo_allocate.set(ar("...اختر العميل المراد تخصيصه..."))
        self.combo_allocate.pack(pady=20)
        
        btn_alloc = ctk.CTkButton(self.main_frame, text=ar("تأكيد التخصيص والنقل"), font=ctk.CTkFont(size=16, weight="bold"), height=50, width=400, command=self.process_allocation)
        btn_alloc.pack(pady=20)

    # ------------------ 3. شاشة طباعة الإيصال ------------------
    def show_receipt_frame(self):
        self.clear_main_frame()
        title = ctk.CTkLabel(self.main_frame, text=ar("طباعة إيصال (PDF)"), font=ctk.CTkFont(size=24, weight="bold"))
        title.pack(pady=(30, 20))
        
        clients = FileManager.get_all_clients()
        if clients:
            self.receipt_map = {ar(c): c for c in clients}
        else:
            self.receipt_map = {ar("لا يوجد عملاء"): ""}

        self.combo_receipt = ctk.CTkComboBox(self.main_frame, values=list(self.receipt_map.keys()), justify="center", width=400, height=40)
        self.combo_receipt.set(ar("...اختر العميل..."))
        self.combo_receipt.pack(pady=20)
        
        self.btn_pdf = ctk.CTkButton(self.main_frame, text=ar("توليد الإيصال (PDF)"), font=ctk.CTkFont(size=16, weight="bold"), height=50, width=400, fg_color="#d9534f", hover_color="#c9302c", command=self.process_receipt)
        self.btn_pdf.pack(pady=20)

    # ================= الأوامر والعمليات =================
    # ملاحظة: نوافذ الرسائل messagebox تستخدم نظام ويندوز الأساسي لذلك لا نستخدم معها دالة ar()
    def process_payment(self):
        selected_ar = self.combo_client.get()
        client = self.client_map.get(selected_ar, "")

        if not client: 
            return messagebox.showwarning("تنبيه", "اختر العميل أولاً.")
        
        try:
            syp = float(self.entry_syp.get())
            usd_val = self.entry_usd.get().strip()
            usd = float(usd_val) if usd_val else 0.0
        except ValueError:
            return messagebox.showerror("خطأ", "المبالغ يجب أن تكون أرقاماً فقط.")

        self.btn_save.configure(state="disabled", text=ar("جاري الترحيل..."))
        self.update()
        success, msg = ExcelHandler.add_payment(client, self.entry_note.get(), self.entry_date.get(), syp, usd)
        self.btn_save.configure(state="normal", text=ar("حفظ وترحيل"))
        messagebox.showinfo("نجاح", msg) if success else messagebox.showerror("خطأ", msg)

    def process_allocation(self):
        selected_ar = self.combo_allocate.get()
        client = self.allocate_map.get(selected_ar, "")

        if not client: 
            return messagebox.showwarning("تنبيه", "اختر العميل أولاً.")

        success, msg = FileManager.move_to_allocated(client)
        if success:
            messagebox.showinfo("نجاح", msg)
            self.show_allocate_frame()
        else:
            messagebox.showerror("خطأ", msg)

    def process_receipt(self):
        selected_ar = self.combo_receipt.get()
        client = self.receipt_map.get(selected_ar, "")

        if not client: 
            return messagebox.showwarning("تنبيه", "اختر العميل أولاً.")

        self.btn_pdf.configure(state="disabled", text=ar("جاري الإنشاء..."))
        self.update()
        success, msg = ExcelHandler.generate_pdf(client)
        self.btn_pdf.configure(state="normal", text=ar("توليد الإيصال (PDF)"))
        messagebox.showinfo("نجاح", msg) if success else messagebox.showerror("خطأ", msg)