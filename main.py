import sys
import os

# 禁用Python緩存文件生成
sys.dont_write_bytecode = True

import tkinter as tk
from tkinter import ttk, messagebox
from auth import login
from database import create_connection, initialize_database, dispatch_scheduled_requirements, has_upcoming_scheduled_requirements
from requirement_manager import RequirementManager
from registration import show_registration_form
import datetime
import time
import threading
import sqlite3

# 檢查資料庫是否存在，如果不存在則從 schema.sql 創建
def initialize_database_from_schema():
    """從 schema.sql 初始化資料庫"""
    db_path = os.path.join(os.path.dirname(__file__), 'users.db')
    schema_path = os.path.join(os.path.dirname(__file__), 'schema.sql')
    
    # 如果資料庫已存在，直接返回
    if os.path.exists(db_path):
        return
        
    print("資料庫不存在，從 schema.sql 創建新資料庫...")
    
    # 讀取 schema.sql 文件
    try:
        with open(schema_path, 'r', encoding='utf-8') as f:
            schema_sql = f.read()
    except Exception as e:
        messagebox.showerror("錯誤", f"無法讀取 schema.sql: {e}")
        sys.exit(1)
        
    # 創建資料庫和表格
    try:
        conn = sqlite3.connect(db_path)
        conn.executescript(schema_sql)
        conn.close()
        print("成功創建資料庫！")
    except Exception as e:
        messagebox.showerror("錯誤", f"創建資料庫失敗: {e}")
        sys.exit(1)

# 初始化資料庫
initialize_database_from_schema()
initialize_database()


# 創建主窗口
root = tk.Tk()
print(f"DEBUG: root window created: {root}")
root.title("需求管理系統")
root.geometry("800x600")  # 調整窗口大小
root.resizable(True, True)


# 定義全局變數
scheduler_running = False
scheduler_thread = None


# 啓動定時任務，檢查並發派預約需求單
def start_global_scheduler():
    """啟動全局定時任務，每10秒檢查一次是否有到期的需求單需要發派"""
    global scheduler_running, scheduler_thread
    
    if scheduler_running:
        return
        
    scheduler_running = True
    
    def check_scheduled_requirements():
        retry_interval = 10  # 正常情況下每10秒檢查一次
        
        while scheduler_running:
            try:
                # 檢查並發派到期的需求單
                conn = create_connection()
                if conn:
                    dispatched_count = dispatch_scheduled_requirements(conn)
                    
                    # 檢查是否有即將到期的需求單，如果有則更頻繁檢查
                    has_upcoming = has_upcoming_scheduled_requirements(conn)
                    conn.close()
                    
                    if dispatched_count > 0:
                        # 使用主線程安全的方式顯示消息（僅當有管理員登入時）
                        root.after(0, lambda: show_dispatch_notification(dispatched_count))
                        
                    # 根據是否有即將到期的需求單動態調整檢查頻率
                    if has_upcoming:
                        retry_interval = 5  # 有即將到期的需求單時，每5秒檢查一次
                    else:
                        retry_interval = 10  # 沒有即將到期的需求單時，每10秒檢查一次
                else:
                    # 無法創建連接時增加重試間隔
                    print("無法創建資料庫連接，稍後重試")
                    retry_interval = min(retry_interval * 2, 60)  # 最多等待1分鐘
            except Exception as e:
                # 發生錯誤時增加重試間隔
                print(f"定時任務執行錯誤: {e}")
                retry_interval = min(retry_interval * 2, 60)  # 最多等待1分鐘
                
            # 等待指定的時間後再次檢查
            for _ in range(retry_interval):
                if not scheduler_running:
                    break
                time.sleep(1)  # 每秒檢查一次是否需要退出
    
    # 在後台線程運行定時任務
    scheduler_thread = threading.Thread(target=check_scheduled_requirements, daemon=True)
    scheduler_thread.start()
    print("全局預約發派調度器已啟動（動態檢查頻率：正常10秒，即將到期時5秒）")


# 顯示發派通知（只有在有管理員登入時才會彈出）
def show_dispatch_notification(count):
    """顯示自動發派通知"""
    global current_app
    if current_app and hasattr(current_app, 'current_user') and current_app.current_user.role == 'admin':
        messagebox.showinfo("自動發派通知", f"已有 {count} 個預約需求單到期並自動發派！")
        # 如果管理員正在查看需求單列表，刷新列表
        if current_app.requirement_manager:
            try:
                # 刷新已發派、預約發派和待審核的需求單列表
                if hasattr(current_app.requirement_manager, 'load_admin_dispatched_requirements'):
                    current_app.requirement_manager.load_admin_dispatched_requirements()
                if hasattr(current_app.requirement_manager, 'load_admin_scheduled_requirements'):
                    current_app.requirement_manager.load_admin_scheduled_requirements()
                if hasattr(current_app.requirement_manager, 'load_admin_reviewing_requirements'):
                    current_app.requirement_manager.load_admin_reviewing_requirements()
            except Exception as e:
                print(f"刷新需求單列表錯誤: {e}")
                pass


# 時間顯示框架
frame_time = ttk.Frame(root)
frame_time.pack(fill=tk.X, padx=10, pady=5)

# 時間標籤
time_label = ttk.Label(frame_time, text="", font=('Arial', 10))
time_label.pack(side=tk.RIGHT)

# 更新時間的函數
def update_time():
    """更新顯示的時間"""
    current_time = datetime.datetime.now().strftime("%Y年%m月%d日 %H:%M:%S")
    time_label.config(text=current_time)
    root.after(1000, update_time)  # 每秒更新一次


# 登入界面
frame_login = ttk.Frame(root, padding=20)
frame_login.pack(pady=50)

ttk.Label(frame_login, text="需求管理系統", font=('Arial', 18, 'bold')).pack(pady=10)

# 用戶名輸入
frame_username = ttk.Frame(frame_login)
frame_username.pack(fill=tk.X, pady=5)
ttk.Label(frame_username, text="使用者名稱:").pack(side=tk.LEFT)
entry_username = ttk.Entry(frame_username)
entry_username.pack(side=tk.RIGHT, padx=5)

# 密碼輸入
frame_password = ttk.Frame(frame_login)
frame_password.pack(fill=tk.X, pady=5)
ttk.Label(frame_password, text="密碼:").pack(side=tk.LEFT)
entry_password = ttk.Entry(frame_password, show="*")
entry_password.pack(side=tk.RIGHT, padx=5)

# 按鈕框架
frame_buttons = ttk.Frame(frame_login)
frame_buttons.pack(pady=10)

# 登入按鈕
button_login = ttk.Button(frame_buttons, text="登入")
button_login.pack(side=tk.LEFT, padx=(0, 10))

# 註冊按鈕
button_register = ttk.Button(frame_buttons, text="註冊新帳號")
button_register.pack(side=tk.LEFT)


# 主要內容區域 (登入後顯示)
frame_main = ttk.Frame(root)

# 功能區域 - 這裡會放置需求管理系統的功能界面（包含個人資料）
frame_content = ttk.Frame(frame_main, padding=10)
frame_content.pack(fill=tk.BOTH, expand=True)

# 保留這些標籤以便兼容性，但它們不會被顯示
label_info_username = ttk.Label(frame_main, text="")
label_info_name = ttk.Label(frame_main, text="")
label_info_email = ttk.Label(frame_main, text="")
label_info_role = ttk.Label(frame_main, text="")
button_logout = ttk.Button(frame_main, text="登出")


class RequirementApp:
    def __init__(self, root, current_user):
        try:
            self.root = root
            self.current_user = current_user
            self.admin_frame = None
            self.staff_frame = None
            self.requirement_manager = None

            if current_user.role == 'admin':
                print("DEBUG: Calling setup_admin_interface")
                self.setup_admin_interface()
                print("DEBUG: Finished setup_admin_interface")
            else:
                print("DEBUG: Calling setup_staff_interface")
                self.setup_staff_interface()
                print("DEBUG: Finished setup_staff_interface")
            print("DEBUG: RequirementApp initialized successfully")
        except Exception as e:
            import traceback
            print(f"ERROR in RequirementApp.__init__: {e}")
            print(traceback.format_exc())
            messagebox.showerror("嚴重錯誤", f"應用程式初始化失敗: {e}")
            root.destroy() # 發生錯誤時銷毀root窗口

    def setup_admin_interface(self):
        """系統管理員界面"""
        try:
            print("DEBUG: Entering setup_admin_interface")
            # 使用需求單管理器設置界面，將界面放在內容區域
            self.requirement_manager = RequirementManager(frame_content, self.current_user)
            self.admin_frame = self.requirement_manager.setup_admin_interface()
            print("DEBUG: Exiting setup_admin_interface")
        except Exception as e:
            import traceback
            print(f"ERROR in setup_admin_interface: {e}")
            print(traceback.format_exc())
            messagebox.showerror("嚴重錯誤", f"管理員界面設定失敗: {e}")
            self.root.destroy()

    def setup_staff_interface(self):
        """一般員工界面"""
        try:
            print("DEBUG: Entering setup_staff_interface")
            # 確保 current_user 是有效的 User 對象，並且有 id 屬性
            if self.current_user is None or not hasattr(self.current_user, 'id') or self.current_user.id is None:
                messagebox.showerror("錯誤", "無法獲取用戶信息，請重新登錄")
                # 改為觸發登出流程，而不是讓整個應用程式崩潰
                if hasattr(self.root, 'event_generate'): # Ensure root is still valid
                    self.root.event_generate("<<LogoutConfirmed>>")
                return # 提前返回，不繼續執行後續的界面設定
            
            print(f"設置員工界面，用戶ID: {self.current_user.id}, 用戶名: {self.current_user.username}")
                
            # 使用需求單管理器設置界面
            self.requirement_manager = RequirementManager(frame_content, self.current_user)
            self.staff_frame = self.requirement_manager.setup_staff_interface()
            print("DEBUG: Exiting setup_staff_interface")
            
        except Exception as e:
            import traceback
            print(f"ERROR in setup_staff_interface: {e}")
            print(traceback.format_exc())
            messagebox.showerror("嚴重錯誤", f"員工界面設定失敗: {e}")
            # 考慮是否真的要 destroy root，或者也可以嘗試登出
            if hasattr(self.root, 'event_generate'):
                self.root.event_generate("<<LogoutConfirmed>>") # 嘗試登出回到登入頁面
            else:
                if self.root and hasattr(self.root, 'destroy'):
                    self.root.destroy() # 最後手段

    def close_interface(self):
        """關閉當前介面"""
        try:
            if self.requirement_manager:
                self.requirement_manager.close()
                self.requirement_manager = None
            
            # 清理管理員界面
            if self.admin_frame:
                self.admin_frame.pack_forget()
                self.admin_frame = None
                
            # 清理員工界面
            if self.staff_frame:
                self.staff_frame.pack_forget()
                self.staff_frame = None
            
        except Exception as e:
            print(f"關閉介面時發生錯誤: {e}")
            # 即使發生錯誤也要確保資源被清理
            pass


# 全局變數存儲當前應用程式實例
current_app = None


def perform_login():
    global current_app
    print(f"DEBUG: perform_login called. current_app is {current_app}")
    if current_app:
        # 如果已經登入，則不執行任何操作
        return
    
    username = entry_username.get()
    password = entry_password.get()

    result = login(username, password)

    if result["success"]:
        frame_login.pack_forget()

        # 顯示使用者資訊
        user = result["user_info"]
        label_info_username.config(text=f"使用者名稱: {user.username}")
        label_info_name.config(text=f"姓名: {user.name}")
        label_info_email.config(text=f"電子郵件: {user.email}")

        # 根據角色顯示不同界面
        if user.role == 'admin':
            label_info_role.config(text=f"角色: 系統管理員")
        else:
            label_info_role.config(text=f"角色: 一般員工")

        frame_main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 初始化主應用程式
        current_app = RequirementApp(root, user)
    else:
        messagebox.showerror("登入失敗", result["message"])


def perform_logout():
    """執行登出操作"""
    # 詢問用戶是否確定要登出
    confirm = messagebox.askyesno("確認登出", "您確定要登出系統嗎？")
    
    if confirm:
        perform_logout_confirmed()

def perform_logout_confirmed():
    """執行已確認的登出操作（不再詢問確認）"""
    try:
        # 關閉當前應用程式介面
        global current_app
        if current_app:
            current_app.close_interface()
            current_app = None
            
        # 清理可能存在的toplevel窗口
        for widget in root.winfo_children():
            if isinstance(widget, tk.Toplevel):
                try:
                    widget.destroy()
                except:
                    pass

        # 清空使用者資訊
        try:
            label_info_username.config(text="")
            label_info_name.config(text="")
            label_info_email.config(text="")
            label_info_role.config(text="")
        except:
            pass
        
        # 隱藏主要內容框架
        try:
            frame_main.pack_forget()
        except:
            pass
        
        # 清空登入表單
        try:
            entry_username.delete(0, tk.END)
            entry_password.delete(0, tk.END)
        except:
            pass
        
        # 顯示登入框架
        try:
            frame_login.pack(pady=50)
        except:
            pass
            
    except Exception as e:
        print(f"登出時發生錯誤: {e}")
        # 即使發生錯誤也要顯示登入界面
        try:
            frame_main.pack_forget()
            frame_login.pack(pady=50)
        except:
            pass


def perform_registration():
    """執行註冊操作"""
    # 呼叫註冊表單模組
    show_registration_form(root)
    # 註冊後自動聚焦於使用者名稱欄位，以便用戶登入
    entry_username.focus_set()


# 設置登入和登出按鈕命令
button_login.config(command=perform_login)
button_logout.config(command=perform_logout)
button_register.config(command=perform_registration)

# 綁定Enter鍵到登入
root.bind("<Return>", lambda event: perform_login())

# 綁定登出事件
root.bind("<<LogoutConfirmed>>", lambda event: perform_logout_confirmed())

# 啟動時間更新
update_time()

# 在應用關閉時停止定時任務
def on_closing():
    """應用關閉時的處理"""
    global scheduler_running, scheduler_thread
    
    # 詢問用戶是否確認關閉應用
    if messagebox.askokcancel("關閉程式", "確定要關閉需求管理系統嗎?"):
        # 停止全局調度器
        scheduler_running = False
        
        # 如果當前有用戶登入，則先登出
        global current_app
        if current_app:
            # 只清理資源，不詢問用戶
            if current_app.requirement_manager:
                current_app.requirement_manager.close()
            current_app = None
            
        # 等待調度器線程結束
        if scheduler_thread and scheduler_thread.is_alive():
            try:
                scheduler_thread.join(timeout=1)  # 最多等待1秒
            except:
                pass
                
        # 關閉應用
        root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)

# 啟動全局定時任務
start_global_scheduler()

print("DEBUG: About to start root.mainloop()")

# 啟動主迴圈
if __name__ == "__main__":
    root.mainloop()
    print("DEBUG: root.mainloop() finished")
else:
    print("DEBUG: Not in __main__, root.mainloop() will not be called directly here.")