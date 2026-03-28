import xlwings as xw
import time

def test_excel_connection():
    print("جاري محاولة فتح برنامج Excel...")
    
    try:
        # 1. فتح تطبيق إكسل وجعله مرئياً (Visible=True) لتراه بعينك
        app = xw.App(visible=True, add_book=False)
        
        # 2. إنشاء كتاب عمل جديد (Workbook)
        wb = app.books.add()
        sheet = wb.sheets[0]
        
        # 3. الكتابة في الخلية A1
        print("جاري الكتابة في الخلية A1...")
        sheet.range('A1').value = 'مرحباً من بايثون!'
        sheet.range('A1').color = (0, 255, 0)  # تلوين الخلية بالأخضر
        
        # انتظر قليلاً لتشاهد التغيير
        time.sleep(2)
        
        # 4. قراءة القيمة من الخلية للتأكد من نجاح العملية
        value = sheet.range('A1').value
        print(f"تمت قراءة القيمة من إكسل: {value}")
        
        if value == 'مرحباً من بايثون!':
            print("\n✅ تهانينا! xlwings تعمل بشكل ممتاز مع نسخة Excel المثبتة.")
        else:
            print("\n❌ هناك مشكلة في قراءة البيانات.")

    except Exception as e:
        print(f"\n❌ حدث خطأ أثناء الاتصال بإكسل: {e}")
        print("تأكد من أن إكسل مفعل ولا تظهر فيه رسائل تمنع التعديل.")

    finally:
        # اترك إكسل مفتوحاً لكي تراه، أو يمكنك إغلاقه بفك التعليق عن السطرين التاليين:
        # wb.close()
        # app.quit()
        print("\nتم الانتهاء من فحص الاتصال.")

if __name__ == "__main__":
    test_excel_connection()