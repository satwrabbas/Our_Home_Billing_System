import shutil
import config
from datetime import datetime

class FileManager:
    @staticmethod
    def setup_directories():
        for folder in config.FOLDERS:
            folder.mkdir(parents=True, exist_ok=True)

    @staticmethod
    def create_client_file(client_name):
        try:
            target_path = config.UNALLOCATED_DIR / f"{client_name}.xlsx"
            if target_path.exists():
                return False, "العميل موجود مسبقاً!"
            
            if not config.UNALLOCATED_TEMPLATE.exists():
                return False, "ملف القالب غير موجود في مجلد Templates"
            
            shutil.copy2(config.UNALLOCATED_TEMPLATE, target_path)
            return True, f"تم إنشاء ملف العميل: {client_name}"
        except Exception as e:
            return False, str(e)

    @staticmethod
    def get_all_clients():
        clients = []
        for file in config.UNALLOCATED_DIR.glob("*.xlsx"):
            if not file.name.startswith("~"):
                clients.append(file.stem)
        return sorted(clients)

    @staticmethod
    def backup_file(file_path):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = config.BACKUPS_DIR / f"{file_path.stem}_BKP_{timestamp}.xlsx"
        shutil.copy2(file_path, backup_path)