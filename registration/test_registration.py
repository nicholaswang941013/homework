import tkinter as tk
import sys
import os

# 添加父目錄到系統路徑
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 導入註冊模組
from registration import show_registration_form
from database import initialize_database

def test_registration():
    """測試註冊功能"""
    # 確保資料庫已初始化
    initialize_database()
    
    # 創建一個臨時的根視窗
    root = tk.Tk()
    root.title("註冊測試")
    root.geometry("300x200")
    
    # 顯示一個說明標籤
    tk.Label(
        root, 
        text="點擊按鈕開始測試註冊功能", 
        font=('Arial', 12)
    ).pack(pady=20)
    
    # 創建測試按鈕
    tk.Button(
        root,
        text="開始測試註冊",
        command=lambda: show_registration_form(root)
    ).pack(pady=10)
    
    # 退出按鈕
    tk.Button(
        root,
        text="退出測試",
        command=root.destroy
    ).pack(pady=10)
    
    # 啟動主循環
    root.mainloop()

if __name__ == "__main__":
    test_registration() 