# config.py
from pathlib import Path

# المسار الأساسي للنظام
BASE_DIR = Path("Real_Estate_System")

# مسارات المجلدات الفرعية
DB_DIR = BASE_DIR / "Database"
TEMPLATES_DIR = DB_DIR / "Templates"
UNALLOCATED_DIR = DB_DIR / "Unallocated"
ALLOCATED_DIR = DB_DIR / "Allocated"
BACKUPS_DIR = BASE_DIR / "Backups"
RECEIPTS_DIR = BASE_DIR / "Receipts_PDF"

# قائمة بكل المجلدات لتسهيل إنشائها
FOLDERS =[TEMPLATES_DIR, UNALLOCATED_DIR, ALLOCATED_DIR, BACKUPS_DIR, RECEIPTS_DIR]