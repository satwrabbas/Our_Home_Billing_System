# excel_handler.py
import xlwings as xw
from datetime import datetime
import config
from file_manager import FileManager

class ExcelHandler:
    @staticmethod
    def add_payment(client_name, installment_name, date_str, amount_syp, amount_usd=0):
        try:
            file_path, status = FileManager.get_client_file(client_name)
            FileManager.backup_file(file_path)
            
            app = xw.App(visible=False)
            wb = app.books.open(file_path)
            sheet1 = wb.sheets["ورقة1"]

            # البحث عن أول صف فارغ
            last_row = 18
            while sheet1.range(f'I{last_row}').value is not None:
                last_row += 1

            # كتابة الدفعة
            sheet1.range(f'A{last_row}').value = installment_name
            sheet1.range(f'D{last_row}').value = amount_syp
            sheet1.range(f'E{last_row}').value = amount_usd
            sheet1.range(f'I{last_row}').value = date_str
            
            # تحديث الإيصال
            sheet3 = wb.sheets["ورقة3"]
            sheet3.range('B6').value = client_name

            wb.save()
            wb.close()
            app.quit()
            return True, "تمت إضافة الدفعة بنجاح وتحديث الإيصال."

        except PermissionError:
            return False, "خطأ: ملف الإكسل مفتوح حالياً. يرجى إغلاقه."
        except Exception as e:
            return False, f"حدث خطأ غير متوقع: {str(e)}"

    @staticmethod
    def generate_pdf(client_name):
        try:
            file_path, status = FileManager.get_client_file(client_name)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            pdf_name = f"إيصال_{client_name}_{timestamp}.pdf"
            pdf_path = config.RECEIPTS_DIR / pdf_name
            
            app = xw.App(visible=False)
            wb = app.books.open(file_path)
            
            receipt_sheet = wb.sheets["ورقة3"]
            receipt_sheet.api.ExportAsFixedFormat(0, str(pdf_path))
            
            wb.close()
            app.quit()
            return True, f"تم إنشاء الإيصال بنجاح!\nتجد الملف في مجلد Receipts_PDF\nباسم: {pdf_name}"
            
        except PermissionError:
            return False, "خطأ: ملف الإكسل مفتوح حالياً. يرجى إغلاقه."
        except Exception as e:
            return False, f"حدث خطأ أثناء إنشاء الـ PDF: {str(e)}"