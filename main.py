# main.py
from file_manager import FileManager
from gui_app import RealEstateApp

if __name__ == "__main__":
    # 1. تهيئة المجلدات عند بدء البرنامج
    FileManager.setup_directories()
    
    # 2. تشغيل واجهة المستخدم
    app = RealEstateApp()
    app.mainloop()