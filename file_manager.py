# file_manager.py
import shutil
from datetime import datetime
import config

class FileManager:
    @staticmethod
    def setup_directories():
        """تهيئة وإنشاء المجلدات إذا لم تكن موجودة"""
        for folder in config.FOLDERS:
            folder.mkdir(parents=True, exist_ok=True)

    @staticmethod
    def get_all_clients():
        """جلب أسماء جميع العملاء من المجلدين"""
        clients = []
        for folder in[config.UNALLOCATED_DIR, config.ALLOCATED_DIR]:
            for file in folder.glob("*.xlsx"):
                if not file.name.startswith("~"):
                    clients.append(file.stem)
        return sorted(clients)

    @staticmethod
    def get_unallocated_clients():
        """جلب أسماء العملاء غير المخصصين فقط"""
        clients =[]
        for file in config.UNALLOCATED_DIR.glob("*.xlsx"):
            if not file.name.startswith("~"):
                clients.append(file.stem)
        return sorted(clients)

    @staticmethod
    def get_client_file(client_name):
        """تحديد مسار ملف العميل وحالته"""
        unallocated_path = config.UNALLOCATED_DIR / f"{client_name}.xlsx"
        allocated_path = config.ALLOCATED_DIR / f"{client_name}.xlsx"

        if allocated_path.exists():
            return allocated_path, "Allocated"
        elif unallocated_path.exists():
            return unallocated_path, "Unallocated"
        raise FileNotFoundError(f"لم يتم العثور على ملف العميل: {client_name}")

    @staticmethod
    def backup_file(file_path):
        """أخذ نسخة احتياطية من الملف"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"{file_path.stem}_backup_{timestamp}{file_path.suffix}"
        backup_path = config.BACKUPS_DIR / backup_name
        shutil.copy2(file_path, backup_path)

    @staticmethod
    def move_to_allocated(client_name):
        """نقل ملف العميل إلى مجلد المتخصص"""
        file_path, status = FileManager.get_client_file(client_name)
        if status == "Allocated":
            return False, "هذا العميل مخصص بالفعل!"
        
        new_path = config.ALLOCATED_DIR / f"{client_name}.xlsx"
        shutil.move(str(file_path), str(new_path))
        return True, f"تم نقل العميل {client_name} بنجاح إلى الشقق المتخصصة."