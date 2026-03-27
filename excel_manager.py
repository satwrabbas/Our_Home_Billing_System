import xlwings as xw
import shutil
from pathlib import Path
from datetime import datetime

class ExcelManager:
    def __init__(self):
        self.base_dir = Path.cwd()
        self.unallocated_dir = self.base_dir / "شقق لاحقة التخصص"
        self.allocated_dir = self.base_dir / "شقق متخصصة"
        self.backup_dir = self.base_dir / "Backups"

        for folder in [self.unallocated_dir, self.allocated_dir, self.backup_dir]:
            folder.mkdir(parents=True, exist_ok=True)

    # -------- الدوال الجديدة لجلب أسماء العملاء تلقائياً -------- #
    def get_unallocated_clients(self):
        """جلب أسماء العملاء غير المتخصصين (بدون صيغة xlsx)"""
        files = self.unallocated_dir.glob("*.xlsx")
        # تجاهل الملفات المؤقتة المفتوحة التي تبدأ بـ ~$
        return [f.stem for f in files if not f.name.startswith('~$')]

    def get_allocated_clients(self):
        """جلب أسماء العملاء المتخصصين"""
        files = self.allocated_dir.glob("*.xlsx")
        return[f.stem for f in files if not f.name.startswith('~$')]

    def get_all_clients(self):
        """جلب جميع العملاء (لإتاحة الدفع لأي عميل)"""
        return self.get_unallocated_clients() + self.get_allocated_clients()
    # ------------------------------------------------------------ #

    def _create_backup(self, file_path: Path):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"{file_path.stem}_{timestamp}{file_path.suffix}"
        backup_path = self.backup_dir / backup_name
        shutil.copy2(file_path, backup_path)
        return backup_path

    def _find_client_file(self, client_name: str):
        file_name = f"{client_name}.xlsx"
        
        unallocated_path = self.unallocated_dir / file_name
        if unallocated_path.exists():
            return unallocated_path, False
        
        allocated_path = self.allocated_dir / file_name
        if allocated_path.exists():
            return allocated_path, True

        raise FileNotFoundError(f"عذراً، لم يتم العثور على ملف العميل: {client_name}")

    def add_payment(self, client_name: str, amount: float, date: str, notes: str):
        try:
            file_path, _ = self._find_client_file(client_name)
        except Exception as e:
            return False, str(e)

        self._create_backup(file_path)
        app = xw.App(visible=False)
        try:
            wb = app.books.open(file_path)
            sheet = wb.sheets['ورقة1'] 

            last_row = sheet.range('I' + str(sheet.cells.last_cell.row)).end('up').row
            new_row = last_row + 1

            sheet.range(f'I{new_row}').value = date
            sheet.range(f'J{new_row}').value = amount
            sheet.range(f'C{new_row}').value = notes

            wb.save()
            return True, f"تمت إضافة الدفعة بنجاح في الصف {new_row}"
        except Exception as e:
            return False, f"حدث خطأ أثناء الكتابة في الملف: {str(e)}"
        finally:
            if 'wb' in locals():
                wb.close()
            app.quit()

    def allocate_apartment(self, client_name: str, area: float, floor_factor: float, direction_factor: float):
        try:
            file_path, is_allocated = self._find_client_file(client_name)
        except Exception as e:
            return False, str(e)

        if is_allocated:
            return False, "هذا العميل مخصص بالفعل!"

        self._create_backup(file_path)
        app = xw.App(visible=False)
        try:
            wb = app.books.open(file_path)
            sheet2 = wb.sheets['ورقة2']

            sheet2.range('B1').value = area
            sheet2.range('L6').value = floor_factor
            
            wb.save()
            wb.close()
            
            new_path = self.allocated_dir / file_path.name
            shutil.move(str(file_path), str(new_path))

            return True, "تم تخصيص الشقة ونقل الملف بنجاح!"
        except Exception as e:
            return False, f"حدث خطأ أثناء التخصيص: {str(e)}"
        finally:
            if len(app.books) == 0:
                app.quit()