import customtkinter as ctk
from tkinter import messagebox
from datetime import datetime
from excel_manager import ExcelManager  # استدعاء العقل الذي بنيناه سابقاً

# إعدادات المظهر العام للتطبيق
ctk.set_appearance_mode("Dark")  # يمكن تغييره إلى "Light" أو "System"
ctk.set_default_color_theme("blue")  # الألوان الأساسية للأزرار

class RealEstateApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # إعدادات النافذة الرئيسية
        self.title("نظام إدارة العقارات والمبيعات الذكي")
        self.geometry("850x600")
        
        # تهيئة قاعدة البيانات (العقل)
        self.db = ExcelManager()

        # تقسيم الشاشة إلى قسمين: قائمة جانبية وشاشة رئيسية
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # ==================== القائمة الجانبية (Sidebar) ====================
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)  # لدفع الأزرار السفلية لأسفل

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="بيتنا العقارية\nOur Home", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 30))

        self.btn_add_payment = ctk.CTkButton(self.sidebar_frame, text="إضافة دفعة مالية", command=self.show_payment_frame)
        self.btn_add_payment.grid(row=1, column=0, padx=20, pady=10)

        self.btn_allocate = ctk.CTkButton(self.sidebar_frame, text="تخصيص شقة", command=self.show_allocate_frame)
        self.btn_allocate.grid(row=2, column=0, padx=20, pady=10)

        # ==================== الشاشة الرئيسية (Main Frame) ====================
        self.main_frame = ctk.CTkFrame(self, corner_radius=15)
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)

        # إعداد واجهة الدفعات كواجهة افتراضية
        self.setup_payment_ui()

    def setup_payment_ui(self):
        """بناء عناصر شاشة إضافة الدفعة"""
        # تنظيف الشاشة الرئيسية أولاً
        for widget in self.main_frame.winfo_children():
            widget.destroy()

        # عنوان الشاشة
        title = ctk.CTkLabel(self.main_frame, text="تسجيل دفعة مالية جديدة", font=ctk.CTkFont(size=24, weight="bold"))
        title.pack(pady=(30, 20))

        # حقل اسم العميل
        self.entry_client = ctk.CTkEntry(self.main_frame, placeholder_text="اسم العميل (مثال: خولة محمد)", justify="right", width=400, height=40)
        self.entry_client.pack(pady=10)

        # حقل رقم القسط أو الملاحظة
        self.entry_note = ctk.CTkEntry(self.main_frame, placeholder_text="البيان (مثال: القسط 34 أو دفعة كاش)", justify="right", width=400, height=40)
        self.entry_note.pack(pady=10)

        # حقل التاريخ (مملوء تلقائياً بتاريخ اليوم)
        today_date = datetime.now().strftime("%Y/%m/%d")
        self.entry_date = ctk.CTkEntry(self.main_frame, justify="right", width=400, height=40)
        self.entry_date.insert(0, today_date)
        self.entry_date.pack(pady=10)

        # حقل المبلغ بالليرة السورية
        self.entry_syp = ctk.CTkEntry(self.main_frame, placeholder_text="المبلغ بالليرة السورية (أرقام فقط)", justify="right", width=400, height=40)
        self.entry_syp.pack(pady=10)

        # حقل المبلغ بالدولار
        self.entry_usd = ctk.CTkEntry(self.main_frame, placeholder_text="المبلغ بالدولار (اختياري - ضع 0 إن لم يوجد)", justify="right", width=400, height=40)
        self.entry_usd.pack(pady=10)

        # زر الحفظ
        self.btn_save = ctk.CTkButton(self.main_frame, text="حفظ وترحيل إلى الإكسل", font=ctk.CTkFont(size=16, weight="bold"), height=50, width=400, fg_color="#28a745", hover_color="#218838", command=self.process_payment)
        self.btn_save.pack(pady=(30, 10))

    def setup_allocate_ui(self):
        """بناء عناصر شاشة تخصيص شقة (سنبرمجها لاحقاً)"""
        for widget in self.main_frame.winfo_children():
            widget.destroy()
        title = ctk.CTkLabel(self.main_frame, text="نقل العميل إلى المتخصص", font=ctk.CTkFont(size=24, weight="bold"))
        title.pack(pady=(30, 20))
        
        self.entry_allocate_client = ctk.CTkEntry(self.main_frame, placeholder_text="اسم العميل لنقله", justify="right", width=400, height=40)
        self.entry_allocate_client.pack(pady=20)
        
        btn_confirm = ctk.CTkButton(self.main_frame, text="تأكيد التخصيص والنقل", font=ctk.CTkFont(size=16, weight="bold"), height=50, width=400, command=self.process_allocation)
        btn_confirm.pack(pady=20)

    # ==================== دوال التنقل بين الشاشات ====================
    def show_payment_frame(self):
        self.setup_payment_ui()

    def show_allocate_frame(self):
        self.setup_allocate_ui()

    # ==================== دوال معالجة العمليات (Business Logic) ====================
    def process_payment(self):
        client = self.entry_client.get().strip()
        note = self.entry_note.get().strip()
        date = self.entry_date.get().strip()
        syp_str = self.entry_syp.get().strip()
        usd_str = self.entry_usd.get().strip()

        # 1. التأكد من تعبئة الحقول الأساسية
        if not client or not note or not syp_str:
            messagebox.showwarning("تنبيه", "يرجى تعبئة كافة الحقول الأساسية (الاسم، البيان، والمبلغ).")
            return

        # 2. التأكد من أن المبالغ هي أرقام (Validation)
        try:
            amount_syp = float(syp_str)
            amount_usd = float(usd_str) if usd_str else 0.0
        except ValueError:
            messagebox.showerror("خطأ إدخال", "يرجى كتابة المبالغ كأرقام صحيحة بدون حروف.")
            return

        # 3. التواصل مع العقل (Backend) لإدخال الدفعة
        # تغيير حالة الزر لمنع الضغط المزدوج
        self.btn_save.configure(text="جاري الترحيل...", state="disabled")
        self.update()

        success, message = self.db.add_payment(client, note, date, amount_syp, amount_usd)

        # إعادة الزر لحالته
        self.btn_save.configure(text="حفظ وترحيل إلى الإكسل", state="normal")

        # 4. عرض النتيجة
        if success:
            messagebox.showinfo("نجاح", message)
            # تفريغ الحقول بعد النجاح
            self.entry_syp.delete(0, 'end')
            self.entry_usd.delete(0, 'end')
            self.entry_note.delete(0, 'end')
        else:
            messagebox.showerror("خطأ", message)

    def process_allocation(self):
        client = self.entry_allocate_client.get().strip()
        if not client:
            messagebox.showwarning("تنبيه", "يرجى كتابة اسم العميل أولاً.")
            return
            
        success, message = self.db.allocate_apartment(client)
        if success:
            messagebox.showinfo("نجاح", message)
            self.entry_allocate_client.delete(0, 'end')
        else:
            messagebox.showerror("خطأ", message)

if __name__ == "__main__":
    app = RealEstateApp()
    app.mainloop()