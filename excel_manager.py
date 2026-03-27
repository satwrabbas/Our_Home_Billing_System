import xlwings as xw
import shutil
from pathlib import Path
from datetime import datetime

class ExcelManager:
    def __init__(self, base_dir="Real_Estate_System"):
        self.base_dir = Path(base_dir)
        self.templates_dir = self.base_dir / "Database" / "Templates"
        self.unallocated_dir = self.base_dir / "Database" / "Unallocated"
        self.allocated_dir = self.base_dir / "Database" / "Allocated"
        self.backups_dir = self.base_dir / "Backups"
        self.receipts_dir = self.base_dir / "Receipts_PDF"

        for folder in[self.templates_dir, self.unallocated_dir, self.allocated_dir, self.backups_dir, self.receipts_dir]:
            folder.mkdir(parents=True, exist_ok=True)

    def get_all_clients(self):
        clients = []
        for folder in[self.unallocated_dir, self.allocated_dir]:
            for file in folder.glob("*.xlsx"):
                if not file.name.startswith("~"):
                    clients.append(file.stem)
        return sorted(clients)

    def get_unallocated_clients(self):
        clients =[]
        for file in self.unallocated_dir.glob("*.xlsx"):
            if not file.name.startswith("~"):
                clients.append(file.stem)
        return sorted(clients)

    def _get_client_file(self, client_name):
        unallocated_path = self.unallocated_dir / f"{client_name}.xlsx"
        allocated_path = self.allocated_dir / f"{client_name}.xlsx"

        if allocated_path.exists():
            return allocated_path, "Allocated"
        elif unallocated_path.exists():
            return unallocated_path, "Unallocated"
        else:
            raise FileNotFoundError(f"لم يتم العثور على ملف العميل: {client_name}")

    def _backup_file(self, file_path):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"{file_path.stem}_backup_{timestamp}{file_path.suffix}"
        backup_path = self.backups_dir / backup_name
        shutil.copy2(file_path, backup_path)

    def add_payment(self, client_name, installment_name, date_str, amount_syp, amount_usd=0):
        try:
            file_path, status = self._get_client_file(client_name)
            self._backup_file(file_path)
            
            app = xw.App(visible=False)
            wb = app.books.open(file_path)
            sheet1 = wb.sheets["ورقة1"]

            last_row = 18
            while sheet1.range(f'I{last_row}').value is not None:
                last_row += 1

            sheet1.range(f'A{last_row}').value = installment_name
            sheet1.range(f'D{last_row}').value = amount_syp
            sheet1.range(f'E{last_row}').value = amount_usd
            sheet1.range(f'I{last_row}').value = date_str
            
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
        try:
            file_path, status = self._get_client_file(client_name)
            
            if status == "Allocated":
                return False, "هذا العميل مخصص بالفعل!"
            
            new_path = self.allocated_dir / f"{client_name}.xlsx"
            shutil.move(str(file_path), str(new_path))
            
            return True, f"تم نقل العميل {client_name} بنجاح إلى مجلد الشقق المتخصصة."
            
        except Exception as e:
            return False, f"حدث خطأ أثناء التخصيص: {str(e)}"

    def generate_receipt_pdf(self, client_name):
        try:
            file_path, status = self._get_client_file(client_name)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            pdf_name = f"إيصال_{client_name}_{timestamp}.pdf"
            pdf_path = self.receipts_dir / pdf_name
            
            app = xw.App(visible=False)
            wb = app.books.open(file_path)
            
            receipt_sheet = wb.sheets["ورقة3"]
            receipt_sheet.api.ExportAsFixedFormat(0, str(pdf_path))
            
            wb.close()
            app.quit()
            
            return True, f"تم إنشاء الإيصال بنجاح!\nتجد الملف في مجلد Receipts_PDF\nباسم: {pdf_name}"
            
        except PermissionError:
            return False, "خطأ: ملف الإكسل مفتوح حالياً. يرجى إغلاقه أولاً لتمكين الطباعة."
        except Exception as e:
            try:
                wb.close()
                app.quit()
            except:
                pass
            return False, f"حدث خطأ أثناء إنشاء الـ PDF: {str(e)}"

if __name__ == "__main__":
    db = ExcelManager()
    print("ExcelManager is working perfectly!")