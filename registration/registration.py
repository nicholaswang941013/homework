import tkinter as tk
from tkinter import ttk, messagebox
import re
import sys
import os

# 添加父目錄到系統路徑，以便可以導入 database 模組
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from database import add_user, create_connection

class RegistrationForm:
    """註冊表單視窗類"""
    
    def __init__(self, parent):
        """初始化註冊表單
        
        Args:
            parent: 父視窗
        """
        self.parent = parent
        self.register_window = None
        self.success = False
        
    def show(self):
        """顯示註冊表單視窗"""
        # 創建註冊視窗
        self.register_window = tk.Toplevel(self.parent)
        self.register_window.title("註冊新帳號")
        self.register_window.geometry("400x450")
        self.register_window.resizable(False, False)
        self.register_window.grab_set()  # 模態對話框
        
        # 設置視窗圖標和樣式
        # self.register_window.iconbitmap('icon.ico')  # 如果有圖標的話
        
        # 標題
        ttk.Label(
            self.register_window, 
            text="註冊新帳號", 
            font=('Arial', 16, 'bold')
        ).pack(pady=(20, 30))
        
        # 表單框架
        form_frame = ttk.Frame(self.register_window, padding=20)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # 用戶名
        ttk.Label(form_frame, text="使用者名稱:").grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        self.username_var = tk.StringVar()
        username_entry = ttk.Entry(form_frame, textvariable=self.username_var, width=30)
        username_entry.grid(row=0, column=1, sticky=tk.W, pady=(0, 10))
        
        # 密碼
        ttk.Label(form_frame, text="密碼:").grid(row=1, column=0, sticky=tk.W, pady=(0, 10))
        self.password_var = tk.StringVar()
        password_entry = ttk.Entry(form_frame, textvariable=self.password_var, show="*", width=30)
        password_entry.grid(row=1, column=1, sticky=tk.W, pady=(0, 10))
        
        # 確認密碼
        ttk.Label(form_frame, text="確認密碼:").grid(row=2, column=0, sticky=tk.W, pady=(0, 10))
        self.confirm_password_var = tk.StringVar()
        confirm_password_entry = ttk.Entry(form_frame, textvariable=self.confirm_password_var, show="*", width=30)
        confirm_password_entry.grid(row=2, column=1, sticky=tk.W, pady=(0, 10))
        
        # 真實姓名
        ttk.Label(form_frame, text="真實姓名:").grid(row=3, column=0, sticky=tk.W, pady=(0, 10))
        self.name_var = tk.StringVar()
        name_entry = ttk.Entry(form_frame, textvariable=self.name_var, width=30)
        name_entry.grid(row=3, column=1, sticky=tk.W, pady=(0, 10))
        
        # 電子郵件
        ttk.Label(form_frame, text="電子郵件:").grid(row=4, column=0, sticky=tk.W, pady=(0, 10))
        self.email_var = tk.StringVar()
        email_entry = ttk.Entry(form_frame, textvariable=self.email_var, width=30)
        email_entry.grid(row=4, column=1, sticky=tk.W, pady=(0, 10))
        
        # 註冊提示
        ttk.Label(
            form_frame, 
            text="所有欄位皆為必填",
            font=('Arial', 9),
            foreground="gray"
        ).grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=(5, 20))
        
        # 按鈕框架
        button_frame = ttk.Frame(form_frame)
        button_frame.grid(row=6, column=0, columnspan=2, sticky=tk.E)
        
        # 取消按鈕
        ttk.Button(
            button_frame,
            text="取消",
            command=self.register_window.destroy
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        # 註冊按鈕
        ttk.Button(
            button_frame,
            text="註冊",
            command=self.perform_registration
        ).pack(side=tk.LEFT)
        
        # 設置初始焦點
        username_entry.focus_set()
        
        # 綁定 Enter 鍵
        self.register_window.bind("<Return>", lambda event: self.perform_registration())
        
        # 等待視窗關閉
        self.parent.wait_window(self.register_window)
        return self.success
        
    def perform_registration(self):
        """執行註冊操作"""
        # 獲取用戶輸入
        username = self.username_var.get().strip()
        password = self.password_var.get().strip()
        confirm_password = self.confirm_password_var.get().strip()
        name = self.name_var.get().strip()
        email = self.email_var.get().strip()
        
        # 驗證輸入
        if not username or not password or not confirm_password or not name or not email:
            messagebox.showerror("錯誤", "所有欄位都必須填寫", parent=self.register_window)
            return
            
        # 密碼匹配檢查
        if password != confirm_password:
            messagebox.showerror("錯誤", "密碼與確認密碼不匹配", parent=self.register_window)
            return
            
        # 郵箱格式檢查
        if not self.is_valid_email(email):
            messagebox.showerror("錯誤", "電子郵件格式不正確", parent=self.register_window)
            return
            
        # 嘗試註冊
        if self.register_user(username, password, name, email):
            messagebox.showinfo("成功", "註冊成功！請使用新帳號登入。", parent=self.register_window)
            self.success = True
            self.register_window.destroy()
        
    def register_user(self, username, password, name, email):
        """向資料庫註冊新使用者
        
        Args:
            username: 使用者名稱
            password: 密碼
            name: 真實姓名
            email: 電子郵件
            
        Returns:
            bool: 註冊是否成功
        """
        try:
            # 所有註冊的帳號皆為員工角色
            result = add_user(username, password, name, email, role='staff')
            
            if not result:
                messagebox.showerror("錯誤", "註冊失敗，使用者名稱可能已被使用", parent=self.register_window)
                return False
                
            return True
        except Exception as e:
            messagebox.showerror("錯誤", f"註冊過程中發生錯誤: {e}", parent=self.register_window)
            return False
            
    def is_valid_email(self, email):
        """檢查郵箱格式是否合法
        
        Args:
            email: 要檢查的電子郵件地址
            
        Returns:
            bool: 郵箱格式是否合法
        """
        # 簡單的郵箱格式檢查
        pattern = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
        return bool(re.match(pattern, email))


def show_registration_form(parent_window):
    """顯示註冊表單
    
    Args:
        parent_window: 父視窗
        
    Returns:
        bool: 註冊是否成功
    """
    registration_form = RegistrationForm(parent_window)
    return registration_form.show()


# 測試代碼，獨立運行時使用
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # 隱藏主視窗
    show_registration_form(root)
    root.destroy() 