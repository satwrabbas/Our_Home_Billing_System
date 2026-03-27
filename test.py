from excel_manager import ExcelManager

# 1. تشغيل مدير الإكسل (سيقوم بإنشاء المجلدات تلقائياً)
manager = ExcelManager()

# ملاحظة: ضع ملف إكسل باسم "خولة محمد.xlsx" داخل مجلد "شقق لاحقة التخصص" يدوياً قبل تشغيل الكود التالي.

# 2. تجربة إضافة دفعة
success, msg = manager.add_payment(
    client_name="خولة محمد",
    amount=5000000,
    date="2026/05/10",
    notes="دفعة نقدية تحت الحساب"
)
print(msg)

# 3. تجربة تخصيص الشقة (سينقل الملف للمجلد الآخر)
success, msg = manager.allocate_apartment(
    client_name="خولة محمد",
    area=120,
    floor_factor=3,
    direction_factor=1.05
)
print(msg)