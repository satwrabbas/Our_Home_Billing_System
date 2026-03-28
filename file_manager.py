# أضف هذه الدالة إلى نهاية الكلاس في ملف file_manager.py
    @staticmethod
    def create_client_file(client_name):
        """يقوم بإنشاء ملف عميل جديد من القالب الأساسي"""
        try:
            # التحقق من أن العميل غير موجود مسبقاً لمنع الكتابة فوقه
            if (config.UNALLOCATED_DIR / f"{client_name}.xlsx").exists() or \
               (config.ALLOCATED_DIR / f"{client_name}.xlsx").exists():
                return False, "خطأ: عميل بهذا الاسم موجود بالفعل!"

            template_path = config.TEMPLATES_DIR / "Template.xlsx"
            if not template_path.exists():
                return False, "خطأ فادح: ملف القالب Template.xlsx غير موجود!"

            new_client_path = config.UNALLOCATED_DIR / f"{client_name}.xlsx"
            shutil.copy2(template_path, new_client_path)
            return True, f"تم إنشاء ملف للعميل '{client_name}' بنجاح في مجلد 'لاحق التخصص'."

        except Exception as e:
            return False, f"حدث خطأ أثناء إنشاء ملف العميل الجديد: {str(e)}"