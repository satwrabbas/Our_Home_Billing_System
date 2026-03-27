import customtkinter as ctk
from tkinter import messagebox
from datetime import datetime
from excel_manager import ExcelManager

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class RealEstateApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("نظام إدارة العقارات والمبيعات الذكي")
        self.geometry("850x600")
        
        self.db = ExcelManager()

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # ==================== القائمة الجانبية ====================
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(5, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="بيتنا العقارية\nOur Home", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 30))

        self.btn_add_payment = ctk.CTkButton(self.sidebar_frame, text="إضافة دفعة مالية", command=self.show_payment_frame)
        self.btn_add_payment.grid(row=1, column=0, padx=20, pady=10)

        self.btn_allocate = ctk.CTkButton(self.sidebar_frame, text="تخصيص شقة", command=self.show_allocate_frame)
        self.btn_allocate.grid(row=2, column=0, padx=20, pady=10)

        # الزر الجديد لطباعة الإيصال
        self.btn_receipt = ctk.CTkButton(self.sidebar_frame, text="استخراج إيصال (PDF)", fg_color="#d9534f", hover_color="#c9302c", command=self.show_receipt_frame)
        self.btn_receipt.grid(row=3, column=0, padx=20, pady=10)

        # ==================== الشاشة الرئيسية ====================
        self.main_frame = ctk.CTkFrame(self, corner_radius=15)
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)

        self.setup_payment_ui()

    # ------------------ 1. شاشة الدفعات ------------------ #
    def setup_payment_ui(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

        title = ctk.CTkLabel(self.main_frame, text="تسجيل دفعة مالية جديدة", font=ctk.CTkFont(size=24, weight="bold"))
        title.pack(pady=(30, 20))

        clients = self.db.get_all_clients()
        if not clients:
            clients =["لا يوجد ملفات عملاء"]

        self.combo_client = ctk.CTkComboBox(self.main_frame, values=clients, justify="center", width=400, height=40)
        self.combo_client.set("...اختر العميل من القائمة أو ابحث...")
        self.combo_client.pack(pady=10)

        self.entry_note = ctk.CTkEntry(self.main_frame, placeholder_text="البيان (مثال: القسط 34 أو دفعة كاش)", justify="center", width=400, height=40)
        self.entry_note.pack(pady=10)

        today_date = datetime.now().strftime("%Y/%m/%d")
        self.entry_date = ctk.CTkEntry(self.main_frame, justify="center", width=400, height=40)
        self.entry_date.insert(0, today_date)
        self.entry_date.pack(pady=10)

        self.entry_syp = ctk.CTkEntry(self.main_frame, placeholder_text="المبلغ بالليرة السورية (أرقام فقط)", justify="center", width=400, height=40)
        self.entry_syp.pack(pady=10)

        self.entry_usd = ctk.CTkEntry(self.main_frame, placeholder_text="المبلغ بالدولار (اختياري - ضع 0 إن لم يوجد)", justify="center", width=400, height=40)
        self.entry_usd.pack(pady=10)

        self.btn_save = ctk.CTkButton(self.main_frame, text="حفظ وترحيل إلى الإكسل", font=ctk.CTkFont(size=16, weight="bold"), height=50, width=400, fg_color="#28a745", hover_color="#218838", command=self.process_payment)
        self.btn_save.pack(pady=(30, 10))

    # ------------------ 2. شاشة التخصيص ------------------ #
    def setup_allocate_ui(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()
            
        title = ctk.CTkLabel(self.main_frame, text="تخصيص شقة (نقل العميل للمتخصص)", font=ctk.CTkFont(size=24, weight="bold"))
        title.pack(pady=(30, 20))
        
        unallocated_clients = self.db.get_unallocated_clients()
        if not unallocated_clients:
            unallocated_clients = ["لا يوجد عملاء غير مخصصين"]
            
        self.combo_allocate_client = ctk.CTkComboBox(self.main_frame, values=unallocated_clients, justify="center", width=400, height=40)
        self.combo_allocate_client.set("...اختر العميل المراد تخصيص شقته...")
        self.combo_allocate_client.pack(pady=20)
        
        btn_confirm = ctk.CTkButton(self.main_frame, text="تأكيد التخصيص والنقل", font=ctk.CTkFont(size=16, weight="bold"), height=50, width=400, command=self.process_allocation)
        btn_confirm.pack(pady=20)

    # ------------------ 3. شاشة طباعة الإيصال ------------------ #
    def setup_receipt_ui(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()
            
        title = ctk.CTkLabel(self.main_frame, text="طباعة إيصال (PDF)", font=ctk.CTkFont(size=24, weight="bold"))
        title.pack(pady=(30, 20))
        
        clients = self.db.get_all_clients()
        if not clients:
            clients = ["لا يوجد ملفات عملاء"]
            
        self.combo_receipt_client = ctk.CTkComboBox(self.main_frame, values=clients, justify="center", width=400, height=40)
        self.combo_receipt_client.set("...اختر العميل لطباعة الإيصال الأخير...")
        self.combo_receipt_client.pack(pady=20)
        
        self.btn_generate_pdf = ctk.CTkButton(self.main_frame, text="توليد الإيصال (PDF)", font=ctk.CTkFont(size=16, weight="bold"), height=50, width=400, fg_color="#d9534f", hover_color="#c9302c", command=self.process_receipt)
        self.btn_generate_pdf.pack(pady=20)

    # ==================== دوال التنقل ====================
    def show_payment_frame(self):
        self.setup_payment_ui()

    def show_allocate_frame(self):
        self.setup_allocate_ui()
        
    def show_receipt_frame(self):
        self.setup_receipt_ui()

    # ==================== دوال معالجة الأوامر ====================
    def process_payment(self):
        client = self.combo_client.get().strip()
        note = self.entry_note.get().strip()
        date = self.entry_date.get().strip()
        syp_str = self.entry_syp.get().strip()
        usd_str = self.entry_usd.get().strip()

        if client in["...اختر العميل من القائمة أو ابحث...", "لا يوجد ملفات عملاء"] or not client:
            messagebox.showwarning("تنبيه", "يرجى اختيار اسم العميل من القائمة.")
            return

        if not note or not syp_str:
            messagebox.showwarning("تنبيه", "يرجى تعبئة كافة الحقول (البيان، والمبلغ).")
            return

        try:
            amount_syp = float(syp_str)
            amount_usd = float(usd_str) if usd_str else 0.0
        except ValueError:
            messagebox.showerror("خطأ إدخال", "يرجى كتابة المبالغ كأرقام صحيحة بدون حروف.")
            return

        self.btn_save.configure(text="جاري الترحيل...", state="disabled")
        self.update()

        success, message = self.db.add_payment(client, note, date, amount_syp, amount_usd)

        self.btn_save.configure(text="حفظ وترحيل إلى الإكسل", state="normal")

        if success:
            messagebox.showinfo("نجاح", message)
            self.entry_syp.delete(0, 'end')
            self.entry_usd.delete(0, 'end')
            self.entry_note.delete(0, 'end')
        else:
            messagebox.showerror("خطأ", message)

    def process_allocation(self):
        client = self.combo_allocate_client.get().strip()
        
        if client in["...اختر العميل المراد تخصيص شقته...", "لا يوجد عملاء غير مخصصين"] or not client:
            messagebox.showwarning("تنبيه", "يرجى اختيار اسم العميل من القائمة.")
            return
            
        success, message = self.db.allocate_apartment(client)
        if success:
            messagebox.showinfo("نجاح", message)
            self.setup_allocate_ui()
        else:
            messagebox.showerror("خطأ", message)

    def process_receipt(self):
        client = self.combo_receipt_client.get().strip()
        
        if client in["...اختر العميل لطباعة الإيصال الأخير...", "لا يوجد ملفات عملاء"] or not client:
            messagebox.showwarning("تنبيه", "يرجى اختيار اسم العميل من القائمة.")
            return
            
        self.btn_generate_pdf.configure(text="جاري إنشاء الـ PDF...", state="disabled")
        self.update()
        
        success, message = self.db.generate_receipt_pdf(client)
        
        self.btn_generate_pdf.configure(text="توليد الإيصال (PDF)", state="normal")
        
        if success:
            messagebox.showinfo("نجاح", message)
        else:
            messagebox.showerror("خطأ", message)

if __name__ == "__main__":
    app = RealEstateApp()
    app.mainloop()