# excel_handler.py
import xlwings as xw
from datetime import datetime
import config
from file_manager import FileManager
import os # تأكد من إضافة هذه المكتبة

class ExcelHandler:
    @staticmethod
    def add_payment(client_name, installment_name, date_str, amount_syp, amount_usd=0):
        app = None
        try:
            file_path, status = FileManager.get_client_file(client_name)
            FileManager.backup_file(file_path)
            
            # استخدام os.path.abspath فهو أقوى في التعامل مع إكسل وويندوز
            absolute_path = os.path.abspath(file_path)
            
            # ========== التعديل السحري هنا ==========
            # سنجعل الإكسل مرئياً مؤقتاً لنرى المشكلة، ونلغي التنبيهات
            app = xw.App(visible=True, add_book=False)
            app.display_alerts = False  # تجاهل النوافذ المنبثقة
            app.screen_updating = False

            # تجاهل رسائل تحديث الروابط (update_links=False)
            wb = app.books.open(absolute_path, update_links=False)
            # ========================================

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

        except Exception as e:
            if app is not None:
                app.quit()
            return False, f"حدث خطأ غير متوقع في الإكسل:\n{str(e)}"

    @staticmethod
    def generate_pdf(client_name):
        app = None
        try:
            file_path, status = FileManager.get_client_file(client_name)
            absolute_excel_path = os.path.abspath(file_path)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            pdf_name = f"إيصال_{client_name}_{timestamp}.pdf"
            pdf_path = config.RECEIPTS_DIR / pdf_name
            absolute_pdf_path = os.path.abspath(pdf_path)
            
            # التعديل هنا أيضاً
            app = xw.App(visible=True, add_book=False)
            app.display_alerts = False
            
            wb = app.books.open(absolute_excel_path, update_links=False)
            receipt_sheet = wb.sheets["ورقة3"]
            receipt_sheet.api.ExportAsFixedFormat(0, absolute_pdf_path)
            
            wb.close()
            app.quit()
            return True, f"تم إنشاء الإيصال بنجاح!\nتجد الملف في مجلد Receipts_PDF\nباسم: {pdf_name}"
            
        except Exception as e:
            if app is not None:
                app.quit()
            return False, f"حدث خطأ أثناء إنشاء الـ PDF:\n{str(e)}"