import xlwings as xw
import shutil
from pathlib import Path
from datetime import datetime

class ExcelManager:
    def __init__(self, base_dir="Real_Estate_System"):
        # إعداد مسارات قاعدة البيانات الذكية
        self.base_dir = Path(base_dir)
        self.templates_dir = self.base_dir / "Database" / "Templates"
        self.unallocated_dir = self.base_dir / "Database" / "Unallocated"
        self.allocated_dir = self.base_dir / "Database" / "Allocated"
        self.backups_dir = self.base_dir / "Backups"
        self.receipts_dir = self.base_dir / "Receipts_PDF"

        # إنشاء المجلدات إذا لم تكن موجودة مسبقاً
        for folder in[self.templates_dir, self.unallocated_dir, self.allocated_dir, self.backups_dir, self.receipts_dir]:
            folder.mkdir(parents=True, exist_ok=True)

    def _get_client_file(self, client_name):
        # البحث عن ملف العميل في المجلدين (متخصص و لاحق التخصص)
        unallocated_path = self.unallocated_dir / f"{client_name}.xlsx"
        allocated_path = self.allocated_dir / f"{client_name}.xlsx"

        if allocated_path.exists():
            return allocated_path, "Allocated"
        elif unallocated_path.exists():
            return unallocated_path, "Unallocated"
        else:
            raise FileNotFoundError(f"لم يتم العثور على ملف العميل: {client_name}")

    def _backup_file(self, file_path):
        # أخذ نسخة احتياطية قبل أي تعديل
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"{file_path.stem}_backup_{timestamp}{file_path.suffix}"
        backup_path = self.backups_dir / backup_name
        shutil.copy2(file_path, backup_path)

    def add_payment(self, client_name, installment_name, date_str, amount_syp, amount_usd=0):
        """
        دالة إضافة دفعة جديدة للعميل
        """
        try:
            file_path, status = self._get_client_file(client_name)
            self._backup_file(file_path) # حماية البيانات
            
            # فتح الإكسل في الخلفية (مخفي لتسريع العمل ومنع تدخل المستخدم)
            app = xw.App(visible=False)
            wb = app.books.open(file_path)
            sheet1 = wb.sheets["ورقة1"]

            # البحث عن أول صف فارغ في العمود I (عمود التاريخ) ابتداءً من الصف 18
            # قمنا باختيار الصف 18 لأن الصفوف الأولى محجوزة للترويسة
            last_row = 18
            while sheet1.range(f'I{last_row}').value is not None:
                last_row += 1

            # إدخال البيانات في الصف الفارغ الجديد
            sheet1.range(f'A{last_row}').value = installment_name  # القسط / الملاحظة
            sheet1.range(f'D{last_row}').value = amount_syp        # المبلغ بالليرة
            sheet1.range(f'E{last_row}').value = amount_usd        # المبلغ بالدولار
            sheet1.range(f'I{last_row}').value = date_str          # التاريخ
            
            # تحديث اسم العميل في ورقة الطباعة (الإيصال)
            sheet3 = wb.sheets["ورقة3"]
            sheet3.range('B6').value = client_name

            wb.save()
            wb.close()
            app.quit()
            
            return True, "تمت إضافة الدفعة بنجاح وتحديث الإيصال."

        except PermissionError:
            return False, "خطأ: ملف الإكسل الخاص بهذا العميل مفتوح حالياً. يرجى إغلاقه والمحاولة مجدداً."
        except Exception as e:
            return False, f"حدث خطأ غير متوقع: {str(e)}"

    def allocate_apartment(self, client_name):
        """
        دالة التخصص: تنقل العميل من مجلد (لاحق التخصص) إلى (متخصص)
        """
        try:
            file_path, status = self._get_client_file(client_name)
            
            if status == "Allocated":
                return False, "هذا العميل مخصص بالفعل!"
            
            # نقل الملف
            new_path = self.allocated_dir / f"{client_name}.xlsx"
            shutil.move(str(file_path), str(new_path))
            
            return True, f"تم نقل العميل {client_name} بنجاح إلى مجلد الشقق المتخصصة."
            
        except Exception as e:
            return False, f"حدث خطأ أثناء التخصيص: {str(e)}"

# ========================================== #
# منطقة الاختبار (Testing)
# ========================================== #
if __name__ == "__main__":
    db = ExcelManager()
    print("تم إنشاء هيكلية قاعدة البيانات بنجاح!")
    
    # يمكنك وضع ملف إكسل باسم "خولة محمد.xlsx" في مجلد Unallocated لتجربة الكود التالي:
    # success, msg = db.add_payment("خولة محمد", "القسط 34", "2026/03/27", 1000000, 500)
    # print(msg)