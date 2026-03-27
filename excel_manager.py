import xlwings as xw
import shutil
from pathlib import Path
from datetime import datetime

class ExcelManager:
    def __init__(self):
        # إعداد مسارات المجلدات الرئيسية باستخدام pathlib
        self.base_dir = Path.cwd()
        self.unallocated_dir = self.base_dir / "شقق لاحقة التخصص"
        self.allocated_dir = self.base_dir / "شقق متخصصة"
        self.backup_dir = self.base_dir / "Backups"

        # إنشاء المجلدات إذا لم تكن موجودة
        for folder in[self.unallocated_dir, self.allocated_dir, self.backup_dir]:
            folder.mkdir(parents=True, exist_ok=True)

    def _create_backup(self, file_path: Path):
        """إنشاء نسخة احتياطية من الملف قبل تعديله"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"{file_path.stem}_{timestamp}{file_path.suffix}"
        backup_path = self.backup_dir / backup_name
        shutil.copy2(file_path, backup_path)
        return backup_path

    def _find_client_file(self, client_name: str):
        """البحث عن ملف العميل في المجلدين وإرجاع مساره وحالته"""
        file_name = f"{client_name}.xlsx"
        
        # البحث في المجلد غير المتخصص
        unallocated_path = self.unallocated_dir / file_name
        if unallocated_path.exists():
            return unallocated_path, False # False يعني غير متخصص
        
        # البحث في المجلد المتخصص
        allocated_path = self.allocated_dir / file_name
        if allocated_path.exists():
            return allocated_path, True # True يعني متخصص

        raise FileNotFoundError(f"عذراً، لم يتم العثور على ملف العميل: {client_name}")

    def add_payment(self, client_name: str, amount: float, date: str, notes: str):
        """إضافة دفعة جديدة للعميل"""
        file_path, _ = self._find_client_file(client_name)
        
        # أخذ نسخة احتياطية أولاً (Zero-Bug Approach)
        self._create_backup(file_path)

        # فتح الإكسل في الخلفية (بدون إظهار الواجهة للمستخدم)
        app = xw.App(visible=False)
        try:
            wb = app.books.open(file_path)
            sheet = wb.sheets['ورقة1'] # اسم الورقة التي تحتوي الدفعات

            # إيجاد أول صف فارغ في عمود التواريخ (العمود I)
            # نفترض أن الدفعات تبدأ من الصف 10 فما بعد
            last_row = sheet.range('I' + str(sheet.cells.last_cell.row)).end('up').row
            new_row = last_row + 1

            # إدخال البيانات (تعديل الحروف حسب ملفك الفعلي)
            sheet.range(f'I{new_row}').value = date     # التاريخ
            sheet.range(f'J{new_row}').value = amount   # المبلغ
            sheet.range(f'C{new_row}').value = notes    # الملاحظات

            wb.save()
            return True, f"تمت إضافة الدفعة بنجاح في الصف {new_row}"
        except Exception as e:
            return False, f"حدث خطأ أثناء الكتابة في الملف: {str(e)}"
        finally:
            # التأكد من إغلاق الإكسل في كل الحالات لمنع بقائه معلقاً في الذاكرة
            if 'wb' in locals():
                wb.close()
            app.quit()

    def allocate_apartment(self, client_name: str, area: float, floor_factor: float, direction_factor: float):
        """تخصيص الشقة: إدخال البيانات في الورقة 2 ثم نقل الملف"""
        file_path, is_allocated = self._find_client_file(client_name)

        if is_allocated:
            return False, "هذا العميل مخصص بالفعل!"

        self._create_backup(file_path)
        app = xw.App(visible=False)
        try:
            wb = app.books.open(file_path)
            sheet2 = wb.sheets['ورقة2']

            # إدخال ثوابت الشقة في الورقة 2 (حسب تحليلي السابق للملف)
            sheet2.range('B1').value = area              # مساحة الشقة
            sheet2.range('L6').value = floor_factor      # معامل الطابق (أو رقم الطابق)
            # يمكنك إضافة المزيد من الخلايا هنا بناءً على المدخلات المطلوبة

            wb.save()
            wb.close()
            
            # نقل الملف إلى مجلد "شقق متخصصة"
            new_path = self.allocated_dir / file_path.name
            shutil.move(str(file_path), str(new_path))

            return True, "تم تخصيص الشقة ونقل الملف بنجاح!"
        except Exception as e:
            return False, f"حدث خطأ أثناء التخصيص: {str(e)}"
        finally:
            if len(app.books) == 0:
                app.quit()