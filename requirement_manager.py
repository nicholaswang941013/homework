import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import shutil
from database import (create_connection, get_all_staff, create_requirement, 
                    get_user_requirements, get_admin_dispatched_requirements,
                    get_admin_scheduled_requirements, dispatch_scheduled_requirements,
                    cancel_scheduled_requirement, submit_requirement, approve_requirement,
                    reject_requirement, invalidate_requirement, get_admin_requirements_by_staff,
                    get_admin_scheduled_by_staff, delete_requirement, restore_requirement,
                    get_deleted_requirements)
import datetime
import threading
import time
import os


class RequirementManager:
    """需求單管理類"""
    
    def __init__(self, root, current_user):
        try:
            print(f"DEBUG: RequirementManager initializing with user {current_user.username}")
            self.root = root
            self.current_user = current_user
            
            # 如果 current_user 是 User 對象，提取 id
            if hasattr(current_user, 'id'):
                self.user_id = current_user.id
            else:
                # 如果是字典或其他類型
                self.user_id = current_user.get('id') if hasattr(current_user, 'get') else current_user['id']
            
            print(f"DEBUG: User ID set to {self.user_id}")
            
            # 用於存儲打開的頂級窗口
            self.open_windows = []
            
            # 建立管理員界面
            if current_user.role == 'admin':
                self.admin_frame = None
                self.staff_combobox = None
                self.title_entry = None
                self.desc_text = None
                self.priority_var = tk.StringVar(value="normal")
                self.dispatch_method_var = tk.StringVar(value="immediate")
                
                # 時間選擇變數
                self.year_var = tk.StringVar()
                self.month_var = tk.StringVar()
                self.day_var = tk.StringVar()
                self.hour_var = tk.StringVar()
                self.minute_var = tk.StringVar()
                
                # 狀態過濾變數
                self.status_filter_var = tk.StringVar(value="all")
                self.staff_filter_var = tk.StringVar(value="all")
                
                # 預約發派員工過濾變數
                self.scheduled_staff_filter_var = tk.StringVar(value="all")
                
            else:
                # 員工界面變數
                self.staff_frame = None
                self.staff_req_treeview = None
                self.staff_status_filter_var = tk.StringVar(value="all")
                
            print("DEBUG: RequirementManager initialization completed successfully")
            
        except Exception as e:
            import traceback
            print(f"ERROR in RequirementManager.__init__: {e}")
            print(traceback.format_exc())
            raise
    
    def get_connection(self):
        """獲取新的資料庫連接"""
        return create_connection()
    
    def execute_with_connection(self, func, *args, **kwargs):
        """使用新連接執行資料庫操作"""
        conn = self.get_connection()
        if not conn:
            return None
        try:
            result = func(conn, *args, **kwargs)
            return result
        except Exception as e:
            print(f"資料庫操作錯誤: {e}")
            return None
        finally:
            try:
                conn.close()
            except:
                pass
    
    def create_toplevel_window(self, title, geometry="600x500"):
        """創建並追蹤 Toplevel 視窗"""
        window = tk.Toplevel(self.root)
        window.title(title)
        window.geometry(geometry)
        window.resizable(True, True)
        window.grab_set()  # 模態對話框
        
        # 追蹤視窗
        self.open_windows.append(window)
        
        # 當視窗關閉時從追蹤列表中移除
        def on_window_close():
            try:
                if window in self.open_windows:
                    self.open_windows.remove(window)
                window.destroy()
            except:
                pass
        
        window.protocol("WM_DELETE_WINDOW", on_window_close)
        return window
    
    def setup_admin_interface(self):
        """設置管理員派發需求單介面"""
        # 創建主容器
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 左側按鈕區域
        button_frame = ttk.Frame(main_container, padding=10)
        button_frame.pack(side=tk.LEFT, fill=tk.Y)
        
        # 右側內容區域
        self.content_frame = ttk.Frame(main_container, padding=10)
        self.content_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 創建垂直按鈕列表
        self.current_tab = "dispatch"  # 預設顯示發派需求單
        
        # 發派需求單按鈕
        self.btn_dispatch = ttk.Button(
            button_frame, 
            text="發派需求單", 
            width=15,
            command=lambda: self.switch_tab("dispatch")
        )
        self.btn_dispatch.pack(pady=5, fill=tk.X)
        
        # 已發派需求單按鈕
        self.btn_dispatched = ttk.Button(
            button_frame, 
            text="已發派需求單", 
            width=15,
            command=lambda: self.switch_tab("dispatched")
        )
        self.btn_dispatched.pack(pady=5, fill=tk.X)
        
        # 待審核需求單按鈕
        self.btn_reviewing = ttk.Button(
            button_frame, 
            text="待審核需求單", 
            width=15,
            command=lambda: self.switch_tab("reviewing")
        )
        self.btn_reviewing.pack(pady=5, fill=tk.X)
        
        # 預約發派需求單按鈕
        self.btn_scheduled = ttk.Button(
            button_frame, 
            text="預約發派需求單", 
            width=15,
            command=lambda: self.switch_tab("scheduled")
        )
        self.btn_scheduled.pack(pady=5, fill=tk.X)
        
        # 垃圾桶按鈕
        self.btn_trash = ttk.Button(
            button_frame, 
            text="垃圾桶", 
            width=15,
            command=lambda: self.switch_tab("trash")
        )
        self.btn_trash.pack(pady=5, fill=tk.X)
        
        # 個人資料按鈕
        self.btn_profile = ttk.Button(
            button_frame, 
            text="個人資料", 
            width=15,
            command=lambda: self.switch_tab("profile")
        )
        self.btn_profile.pack(pady=5, fill=tk.X)
        
        # 創建各個標籤頁的內容框架
        self.dispatch_tab = ttk.Frame(self.content_frame)
        self.dispatched_tab = ttk.Frame(self.content_frame)
        self.reviewing_tab = ttk.Frame(self.content_frame)
        self.scheduled_tab = ttk.Frame(self.content_frame)
        self.trash_tab = ttk.Frame(self.content_frame)
        self.profile_tab = ttk.Frame(self.content_frame)
        
        # 設置各個標籤頁的內容
        self.setup_dispatch_tab(self.dispatch_tab)
        self.setup_dispatched_tab(self.dispatched_tab)
        self.setup_reviewing_tab(self.reviewing_tab)
        self.setup_scheduled_tab(self.scheduled_tab)
        self.setup_trash_tab(self.trash_tab)
        self.setup_profile_tab(self.profile_tab)
        
        # 預設顯示發派需求單
        self.switch_tab("dispatch")
        
        return main_container
    
    def switch_tab(self, tab_name):
        """切換標籤頁"""
        # 隱藏所有標籤頁
        for tab in [self.dispatch_tab, self.dispatched_tab, self.reviewing_tab, self.scheduled_tab, self.trash_tab, self.profile_tab]:
            tab.pack_forget()
        
        # 重置所有按鈕狀態
        for btn in [self.btn_dispatch, self.btn_dispatched, self.btn_reviewing, self.btn_scheduled, self.btn_trash, self.btn_profile]:
            btn.state(['!pressed'])
        
        # 顯示選中的標籤頁並設置按鈕狀態
        if tab_name == "dispatch":
            self.dispatch_tab.pack(fill=tk.BOTH, expand=True)
            self.btn_dispatch.state(['pressed'])
        elif tab_name == "dispatched":
            self.dispatched_tab.pack(fill=tk.BOTH, expand=True)
            self.btn_dispatched.state(['pressed'])
            self.load_admin_dispatched_requirements()
        elif tab_name == "reviewing":
            self.reviewing_tab.pack(fill=tk.BOTH, expand=True)
            self.btn_reviewing.state(['pressed'])
            self.load_admin_reviewing_requirements()
        elif tab_name == "scheduled":
            self.scheduled_tab.pack(fill=tk.BOTH, expand=True)
            self.btn_scheduled.state(['pressed'])
            self.load_admin_scheduled_requirements()
        elif tab_name == "trash":
            self.trash_tab.pack(fill=tk.BOTH, expand=True)
            self.btn_trash.state(['pressed'])
            self.load_deleted_requirements()
        elif tab_name == "profile":
            self.profile_tab.pack(fill=tk.BOTH, expand=True)
            self.btn_profile.state(['pressed'])
        
        self.current_tab = tab_name
        
    def setup_dispatch_tab(self, parent):
        """設置發派需求單標籤頁"""
        # 發派需求單框架
        self.admin_frame = ttk.LabelFrame(parent, text="發派需求單", padding=10)
        self.admin_frame.pack(pady=10, fill=tk.BOTH)
        
        # 創建管理員界面元素 - 使用網格佈局
        # 指派對象選擇（移到最上面）
        ttk.Label(self.admin_frame, text="指派給:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        # 創建下拉選單來選擇員工
        staffs = self.execute_with_connection(get_all_staff) or []
        staff_list = [f"{staff[1]} (ID:{staff[0]})" for staff in staffs]
        
        self.staff_combobox = ttk.Combobox(self.admin_frame, values=staff_list, width=37)
        self.staff_combobox.grid(row=0, column=1, sticky=tk.W, padx=5, pady=(0, 5))
        
        # 標題輸入（移到第二位，增加上邊距）
        ttk.Label(self.admin_frame, text="標題:").grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        self.title_entry = ttk.Entry(self.admin_frame, width=40)
        self.title_entry.grid(row=1, column=1, sticky=tk.W+tk.E, padx=5, pady=(10, 0))
        
        # 內容輸入
        ttk.Label(self.admin_frame, text="內容:").grid(row=2, column=0, sticky=tk.NW, pady=(10, 0))
        self.desc_text = tk.Text(self.admin_frame, width=40, height=8)
        self.desc_text.grid(row=2, column=1, sticky=tk.W+tk.E, padx=5, pady=(10, 5))
        
        # 緊急程度選擇
        ttk.Label(self.admin_frame, text="緊急程度:").grid(row=3, column=0, sticky=tk.W)
        self.priority_var = tk.StringVar(value="normal")
        
        priority_frame = ttk.Frame(self.admin_frame)
        priority_frame.grid(row=3, column=1, sticky=tk.W, pady=5)
        
        ttk.Radiobutton(
            priority_frame, 
            text="普通", 
            value="normal", 
            variable=self.priority_var
        ).pack(side=tk.LEFT, padx=10)
        
        ttk.Radiobutton(
            priority_frame, 
            text="緊急", 
            value="urgent", 
            variable=self.priority_var
        ).pack(side=tk.LEFT, padx=10)

        # 發派時間選擇
        ttk.Label(self.admin_frame, text="發派方式:").grid(row=4, column=0, sticky=tk.W)
        self.dispatch_method_var = tk.StringVar(value="immediate")
        
        dispatch_method_frame = ttk.Frame(self.admin_frame)
        dispatch_method_frame.grid(row=4, column=1, sticky=tk.W, pady=5)
        
        ttk.Radiobutton(
            dispatch_method_frame,
            text="立即發派",
            value="immediate",
            variable=self.dispatch_method_var,
            command=self.toggle_schedule_frame
        ).pack(side=tk.LEFT, padx=10)
        
        ttk.Radiobutton(
            dispatch_method_frame,
            text="預約發派",
            value="scheduled",
            variable=self.dispatch_method_var,
            command=self.toggle_schedule_frame
        ).pack(side=tk.LEFT, padx=10)
        
        # 預約時間選擇框架（初始隱藏）
        self.schedule_frame = ttk.LabelFrame(self.admin_frame, text="預約發派設定", padding=5)
        self.schedule_frame.grid(row=5, column=0, columnspan=2, sticky=tk.W+tk.E, pady=5)
        self.schedule_frame.grid_remove()  # 初始隱藏
        
        # 日期選擇
        date_frame = ttk.Frame(self.schedule_frame)
        date_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(date_frame, text="日期:").pack(side=tk.LEFT, padx=(0, 5))
        
        # 年份選擇
        current_year = datetime.datetime.now().year
        self.year_var = tk.StringVar(value=str(current_year))
        years = [str(current_year + i) for i in range(0, 5)]  # 當前年份及未來4年
        ttk.Combobox(
            date_frame, 
            textvariable=self.year_var,
            values=years,
            width=6
        ).pack(side=tk.LEFT)
        
        ttk.Label(date_frame, text="年").pack(side=tk.LEFT)
        
        # 月份選擇
        current_month = datetime.datetime.now().month
        self.month_var = tk.StringVar(value=str(current_month))
        months = [str(i) for i in range(1, 13)]
        ttk.Combobox(
            date_frame, 
            textvariable=self.month_var,
            values=months,
            width=4
        ).pack(side=tk.LEFT)
        
        ttk.Label(date_frame, text="月").pack(side=tk.LEFT)
        
        # 日期選擇
        current_day = datetime.datetime.now().day
        self.day_var = tk.StringVar(value=str(current_day))
        days = [str(i) for i in range(1, 32)]
        ttk.Combobox(
            date_frame, 
            textvariable=self.day_var,
            values=days,
            width=4
        ).pack(side=tk.LEFT)
        
        ttk.Label(date_frame, text="日").pack(side=tk.LEFT)
        
        # 時間選擇
        time_frame = ttk.Frame(self.schedule_frame)
        time_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(time_frame, text="時間:").pack(side=tk.LEFT, padx=(0, 5))
        
        # 小時選擇
        current_hour = datetime.datetime.now().hour
        self.hour_var = tk.StringVar(value=str(current_hour))
        hours = [str(i).zfill(2) for i in range(0, 24)]
        ttk.Combobox(
            time_frame, 
            textvariable=self.hour_var,
            values=hours,
            width=4
        ).pack(side=tk.LEFT)
        
        ttk.Label(time_frame, text="時").pack(side=tk.LEFT)
        
        # 分鐘選擇
        self.minute_var = tk.StringVar(value="00")
        minutes = [str(i).zfill(2) for i in range(0, 60, 5)]  # 以5分鐘為間隔
        ttk.Combobox(
            time_frame, 
            textvariable=self.minute_var,
            values=minutes,
            width=4
        ).pack(side=tk.LEFT)
        
        ttk.Label(time_frame, text="分").pack(side=tk.LEFT)

        # 派發按鈕
        ttk.Button(
            self.admin_frame,
            text="派發需求單",
            command=self.create_requirement
        ).grid(row=6, column=1, pady=10, sticky=tk.E)
        
    def setup_dispatched_tab(self, parent):
        """設置已發派需求單標籤頁"""
        # 已發派需求單框架
        frame = ttk.Frame(parent, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 標題
        ttk.Label(
            frame, 
            text="已發派需求單列表", 
            font=("Arial", 12, "bold")
        ).pack(side=tk.TOP, anchor=tk.W, pady=(0, 10))
        
        # 過濾框架
        filter_frame = ttk.Frame(frame)
        filter_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 狀態過濾
        ttk.Label(filter_frame, text="狀態過濾:").pack(side=tk.LEFT, padx=(0, 10))
        
        self.status_filter_var = tk.StringVar(value="all")
        
        statuses = [
            ("全部", "all"),
            ("未完成", "pending"),
            ("待審核", "reviewing"),
            ("已完成", "completed"),
            ("已失效", "invalid")
        ]
        
        for text, value in statuses:
            ttk.Radiobutton(
                filter_frame,
                text=text,
                value=value,
                variable=self.status_filter_var,
                command=self.load_admin_dispatched_requirements
            ).pack(side=tk.LEFT, padx=5)
        
        # 員工過濾
        staff_filter_frame = ttk.Frame(frame)
        staff_filter_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(staff_filter_frame, text="員工過濾:").pack(side=tk.LEFT, padx=(0, 10))
        
        # 獲取所有員工
        staffs = self.execute_with_connection(get_all_staff) or []
        staff_options = [("全部員工", "all")] + [(staff[1], str(staff[0])) for staff in staffs]
        
        # 建立員工過濾下拉選單
        self.staff_filter_var = tk.StringVar(value="all")
        self.staff_filter_combobox = ttk.Combobox(
            staff_filter_frame,
            textvariable=self.staff_filter_var,
            values=[f"{name} ({id})" for name, id in staff_options],
            width=20
        )
        self.staff_filter_combobox.pack(side=tk.LEFT, padx=5)
        self.staff_filter_combobox.current(0)
        
        # 綁定事件
        self.staff_filter_combobox.bind("<<ComboboxSelected>>", 
                                         lambda event: self.load_admin_dispatched_requirements())
        
        # 創建已發派需求單列表
        columns = ("id", "title", "assignee", "status", "priority", "created_at")
        
        self.admin_dispatched_treeview = ttk.Treeview(
            frame, 
            columns=columns,
            show="headings", 
            selectmode="browse"
        )
        
        # 設置列標題
        self.admin_dispatched_treeview.heading("id", text="ID")
        self.admin_dispatched_treeview.heading("title", text="標題")
        self.admin_dispatched_treeview.heading("assignee", text="指派給")
        self.admin_dispatched_treeview.heading("status", text="狀態")
        self.admin_dispatched_treeview.heading("priority", text="緊急程度")
        self.admin_dispatched_treeview.heading("created_at", text="發派時間")
        
        # 設置列寬
        self.admin_dispatched_treeview.column("id", width=50)
        self.admin_dispatched_treeview.column("title", width=200)
        self.admin_dispatched_treeview.column("assignee", width=100)
        self.admin_dispatched_treeview.column("status", width=80)
        self.admin_dispatched_treeview.column("priority", width=80)
        self.admin_dispatched_treeview.column("created_at", width=150)
        
        # 綁定雙擊事件
        self.admin_dispatched_treeview.bind("<Double-1>", self.show_dispatched_details)
        
        # 添加滾動條
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.admin_dispatched_treeview.yview)
        self.admin_dispatched_treeview.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.admin_dispatched_treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 刷新按鈕
        ttk.Button(
            parent,
            text="刷新列表",
            command=self.load_admin_dispatched_requirements
        ).pack(side=tk.BOTTOM, pady=10)
        
        # 載入數據
        self.load_admin_dispatched_requirements()
        
    def setup_reviewing_tab(self, parent):
        """設置待審核需求單標籤頁"""
        frame = ttk.Frame(parent, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 控制區域
        control_frame = ttk.Frame(frame)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 刷新按鈕
        refresh_button = ttk.Button(
            control_frame, 
            text="刷新", 
            command=self.load_admin_reviewing_requirements
        )
        refresh_button.pack(side=tk.RIGHT)
        
        # 說明標籤
        ttk.Label(
            control_frame,
            text="待員工提交審核的需求單列表"
        ).pack(side=tk.LEFT)
        
        # 需求單列表
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # 建立Treeview
        columns = ("id", "title", "assignee", "priority", "created_at")
        self.admin_reviewing_treeview = ttk.Treeview(list_frame, columns=columns, show='headings')
        self.admin_reviewing_treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 配置列
        self.admin_reviewing_treeview.heading("id", text="ID")
        self.admin_reviewing_treeview.heading("title", text="標題")
        self.admin_reviewing_treeview.heading("assignee", text="接收人")
        self.admin_reviewing_treeview.heading("priority", text="緊急程度")
        self.admin_reviewing_treeview.heading("created_at", text="發派時間")
        
        # 設置列寬
        self.admin_reviewing_treeview.column("id", width=50, anchor=tk.CENTER)
        self.admin_reviewing_treeview.column("title", width=250)
        self.admin_reviewing_treeview.column("assignee", width=100, anchor=tk.CENTER)
        self.admin_reviewing_treeview.column("priority", width=80, anchor=tk.CENTER)
        self.admin_reviewing_treeview.column("created_at", width=120, anchor=tk.CENTER)
        
        # 添加垂直滾動條
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.admin_reviewing_treeview.yview)
        self.admin_reviewing_treeview.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 綁定雙擊事件
        self.admin_reviewing_treeview.bind("<Double-1>", self.show_reviewing_requirement_details)
        
        # 載入數據
        self.load_admin_reviewing_requirements()
        
    def load_admin_reviewing_requirements(self):
        """載入管理員待審核的需求單數據"""
        # 清空現有數據
        for item in self.admin_reviewing_treeview.get_children():
            self.admin_reviewing_treeview.delete(item)
            
        # 獲取所有已發派的需求單 (這些應包含完整的15個欄位)
        requirements = self.execute_with_connection(get_admin_dispatched_requirements, self.user_id) or []
        
        # 篩選狀態為「待審核」的需求單
        reviewing_requirements = [req for req in requirements if len(req) == 15 and req[3] == 'reviewing']
        
        # 添加數據到表格
        for req in reviewing_requirements:
            try:
                # 解包完整的15個欄位
                (req_id, title, description, status, priority, created_at, 
                 assigner_name, assigner_id, assignee_name, assignee_id,
                 scheduled_time, comment, completed_at, attachment_path, deleted_at) = req
                
                # 格式化緊急程度
                priority_text = "緊急" if priority == "urgent" else "普通"
                
                # 格式化時間
                date_text = created_at # Default
                if isinstance(created_at, str):
                    try:
                        date_obj = datetime.datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S")
                        date_text = date_obj.strftime("%Y-%m-%d %H:%M")
                    except ValueError:
                        pass # date_text remains original created_at
                
                # Treeview columns: ("id", "title", "assignee", "priority", "created_at")
                item_id_val = self.admin_reviewing_treeview.insert(
                    "", tk.END, 
                    values=(req_id, title, assignee_name, priority_text, date_text)
                )
                
                # 根據優先級設置行顏色
                if priority == 'urgent':
                    self.admin_reviewing_treeview.item(item_id_val, tags=('urgent',))
                
            except Exception as e:
                print(f"處理待審核需求單時發生錯誤: {e}, 數據: {req}")
                import traceback
                print(traceback.format_exc())
                    
        # 設置標籤顏色
        self.admin_reviewing_treeview.tag_configure('urgent', background='#d4edff')
        
    def setup_scheduled_tab(self, parent):
        """設置預約發派需求單標籤頁"""
        # 預約發派需求單框架
        frame = ttk.Frame(parent, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 標題
        ttk.Label(
            frame, 
            text="預約發派需求單列表", 
            font=("Arial", 12, "bold")
        ).pack(side=tk.TOP, anchor=tk.W, pady=(0, 10))
        
        # 員工過濾框架
        staff_filter_frame = ttk.Frame(frame)
        staff_filter_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(staff_filter_frame, text="員工過濾:").pack(side=tk.LEFT, padx=(0, 10))
        
        # 獲取所有員工
        staffs = self.execute_with_connection(get_all_staff) or []
        staff_options = [("全部員工", "all")] + [(staff[1], str(staff[0])) for staff in staffs]
        
        # 建立員工過濾下拉選單
        self.scheduled_staff_filter_var = tk.StringVar(value="all")
        self.scheduled_staff_filter_combobox = ttk.Combobox(
            staff_filter_frame,
            textvariable=self.scheduled_staff_filter_var,
            values=[f"{name} ({id})" for name, id in staff_options],
            width=20
        )
        self.scheduled_staff_filter_combobox.pack(side=tk.LEFT, padx=5)
        self.scheduled_staff_filter_combobox.current(0)
        
        # 綁定事件
        self.scheduled_staff_filter_combobox.bind("<<ComboboxSelected>>", 
                                          lambda event: self.load_admin_scheduled_requirements())
        
        # 創建預約發派需求單列表
        columns = ("id", "title", "assignee", "priority", "scheduled_time")
        
        self.admin_scheduled_treeview = ttk.Treeview(
            frame, 
            columns=columns,
            show="headings", 
            selectmode="browse"
        )
        
        # 設置列標題
        self.admin_scheduled_treeview.heading("id", text="ID")
        self.admin_scheduled_treeview.heading("title", text="標題")
        self.admin_scheduled_treeview.heading("assignee", text="指派給")
        self.admin_scheduled_treeview.heading("priority", text="緊急程度")
        self.admin_scheduled_treeview.heading("scheduled_time", text="預約發派時間")
        
        # 設置列寬
        self.admin_scheduled_treeview.column("id", width=50)
        self.admin_scheduled_treeview.column("title", width=200)
        self.admin_scheduled_treeview.column("assignee", width=100)
        self.admin_scheduled_treeview.column("priority", width=80)
        self.admin_scheduled_treeview.column("scheduled_time", width=150)
        
        # 綁定雙擊事件
        self.admin_scheduled_treeview.bind("<Double-1>", self.show_scheduled_details)
        
        # 添加滾動條
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.admin_scheduled_treeview.yview)
        self.admin_scheduled_treeview.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.admin_scheduled_treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 按鈕框架
        button_frame = ttk.Frame(parent)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        
        # 取消預約按鈕
        ttk.Button(
            button_frame,
            text="取消選中的預約",
            command=self.cancel_scheduled_requirement
        ).pack(side=tk.LEFT, padx=5)
        
        # 刷新按鈕
        ttk.Button(
            button_frame,
            text="刷新列表",
            command=self.load_admin_scheduled_requirements
        ).pack(side=tk.RIGHT, padx=5)
        
        # 載入數據
        self.load_admin_scheduled_requirements()

    def setup_staff_interface(self, parent):
        """設置員工查看需求單介面"""
        # 創建主容器
        main_container = ttk.Frame(parent) # Should use self.root
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 左側按鈕區域
        button_frame = ttk.Frame(main_container, padding=10)
        button_frame.pack(side=tk.LEFT, fill=tk.Y)
        
        # 右側內容區域
        self.content_frame = ttk.Frame(main_container, padding=10)
        self.content_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 創建垂直按鈕列表
        self.current_tab = "requirements"  # 預設顯示需求單列表
        
        # 我的需求單按鈕
        self.btn_requirements = ttk.Button(
            button_frame, 
            text="我的需求單", 
            width=15,
            command=lambda: self.switch_staff_tab("requirements")
        )
        self.btn_requirements.pack(pady=5, fill=tk.X)
        
        # 個人資料按鈕
        self.btn_staff_profile = ttk.Button(
            button_frame, 
            text="個人資料", 
            width=15,
            command=lambda: self.switch_staff_tab("profile")
        )
        self.btn_staff_profile.pack(pady=5, fill=tk.X)
        
        # 創建標籤頁的內容框架
        self.requirements_tab = ttk.Frame(self.content_frame)
        self.staff_profile_tab = ttk.Frame(self.content_frame)
        
        # 設置標籤頁的內容
        self.setup_requirements_tab(self.requirements_tab) # This passes the correct parent to setup_requirements_tab
        self.setup_profile_tab(self.staff_profile_tab)     # This passes the correct parent to setup_profile_tab
        
        # 預設顯示需求單列表
        self.switch_staff_tab("requirements")
        
        return main_container
    
    def switch_staff_tab(self, tab_name):
        """切換員工界面標籤頁"""
        # 隱藏所有標籤頁
        for tab in [self.requirements_tab, self.staff_profile_tab]:
            tab.pack_forget()
        
        # 重置所有按鈕狀態
        for btn in [self.btn_requirements, self.btn_staff_profile]:
            btn.state(['!pressed'])
        
        # 顯示選中的標籤頁並設置按鈕狀態
        if tab_name == "requirements":
            self.requirements_tab.pack(fill=tk.BOTH, expand=True)
            self.btn_requirements.state(['pressed'])
            self.load_user_requirements()
        elif tab_name == "profile":
            self.staff_profile_tab.pack(fill=tk.BOTH, expand=True)
            self.btn_staff_profile.state(['pressed'])
        
        self.current_tab = tab_name
    
    def setup_requirements_tab(self, parent):
        """設置需求單列表標籤頁"""
        # 需求單列表框架
        self.staff_frame = ttk.LabelFrame(parent, text="我收到的需求單", padding=10)
        self.staff_frame.pack(pady=10, fill=tk.BOTH, expand=True)
        
        # 狀態過濾框架
        filter_frame = ttk.Frame(self.staff_frame)
        filter_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(filter_frame, text="狀態過濾:").pack(side=tk.LEFT, padx=(0, 10))
        
        self.staff_status_filter_var = tk.StringVar(value="all")
        
        statuses = [
            ("全部", "all"),
            ("未完成", "pending"),
            ("待審核", "reviewing"),
            ("已完成", "completed"),
            ("已失效", "invalid")
        ]
        
        for text, value in statuses:
            ttk.Radiobutton(
                filter_frame,
                text=text,
                value=value,
                variable=self.staff_status_filter_var,
                command=self.load_user_requirements
            ).pack(side=tk.LEFT, padx=5)
        
        # 創建需求單列表
        columns = ("id", "title", "assigner", "status", "priority", "date")
        self.staff_req_treeview = ttk.Treeview(
            self.staff_frame, 
            columns=columns,
            show="headings",
            selectmode="browse"
        )
        
        # 設置列標題
        self.staff_req_treeview.heading("id", text="編號")
        self.staff_req_treeview.heading("title", text="標題")
        self.staff_req_treeview.heading("assigner", text="指派人")
        self.staff_req_treeview.heading("status", text="狀態")
        self.staff_req_treeview.heading("priority", text="緊急程度")
        self.staff_req_treeview.heading("date", text="派發日期")
        
        # 設置列寬度
        self.staff_req_treeview.column("id", width=50)
        self.staff_req_treeview.column("title", width=200)
        self.staff_req_treeview.column("assigner", width=100)
        self.staff_req_treeview.column("status", width=100)
        self.staff_req_treeview.column("priority", width=100)
        self.staff_req_treeview.column("date", width=120)
        
        # 添加滾動條
        scrollbar = ttk.Scrollbar(self.staff_frame, orient=tk.VERTICAL, command=self.staff_req_treeview.yview)
        self.staff_req_treeview.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.staff_req_treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 綁定雙擊事件以查看詳情
        self.staff_req_treeview.bind("<Double-1>", self.show_requirement_details)
        
        # 載入需求單數據
        self.load_user_requirements()
        
        # 按鈕框架
        button_frame = ttk.Frame(self.staff_frame)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        
        # 提交需求單按鈕
        ttk.Button(
            button_frame,
            text="提交選中的需求單",
            command=self.submit_requirement
        ).pack(side=tk.LEFT, padx=5)
        
        # 刷新按鈕
        ttk.Button(
            button_frame,
            text="刷新列表",
            command=self.load_user_requirements
        ).pack(side=tk.RIGHT, padx=5)
    
    def load_user_requirements(self):
        """載入用戶收到的需求單到列表"""
        # 清空現有數據
        for item in self.staff_req_treeview.get_children():
            self.staff_req_treeview.delete(item)
        
        # 檢查用戶ID是否有效
        if self.user_id is None:
            messagebox.showerror("錯誤", "無法獲取用戶ID，請重新登錄")
            return
            
        # 獲取數據
        requirements = self.execute_with_connection(get_user_requirements, self.user_id) or []
        
        # 獲取狀態過濾條件
        status_filter = self.staff_status_filter_var.get()
        
        if not requirements:
            empty_id = self.staff_req_treeview.insert("", tk.END, values=("", "目前沒有收到任何需求單", "", "", "", ""))
            self.staff_req_treeview.item(empty_id, tags=('empty',))
            self.staff_req_treeview.tag_configure('empty', foreground='gray')
            return
        
        # 添加數據到表格
        for req in requirements:
            if len(req) < 15:
                print(f"警告: 員工需求單資料不完整，跳過: {req}")
                continue

            # Unpack all 15 fields from _get_requirement_select_fields
            (req_id, title, description, status, priority, created_at, 
             assigner_name, assigner_id, 
             assignee_name, assignee_id, 
             scheduled_time, comment, completed_at, attachment_path, deleted_at) = req
            
            if status_filter != "all" and status != status_filter:
                continue
                
            status_text = self.get_status_display_text(status)
            priority_text = "緊急" if priority == "urgent" else "普通"
            
            if isinstance(created_at, str):
                try:
                    date_obj = datetime.datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S")
                    date_text = date_obj.strftime("%Y-%m-%d %H:%M")
                except ValueError:
                    date_text = created_at
            else:
                date_text = created_at
                
            item_id = self.staff_req_treeview.insert(
                "", tk.END, 
                values=(req_id, title, assigner_name, status_text, priority_text, date_text)
            )
            
            if status == 'reviewing':
                self.staff_req_treeview.item(item_id, tags=('reviewing',))
            elif status == 'completed':
                self.staff_req_treeview.item(item_id, tags=('completed',))
            elif status == 'invalid':
                self.staff_req_treeview.item(item_id, tags=('invalid',))
                
        self.staff_req_treeview.tag_configure('reviewing', background='#d4edff')

    def show_requirement_details(self, event):
        """顯示需求單詳情"""
        # 獲取選中的項目
        selected_item = self.staff_req_treeview.selection()
        if not selected_item:
            return
            
        item = self.staff_req_treeview.item(selected_item)
        req_id = item['values'][0]
        
        # 獲取需求單詳情
        # IMPORTANT: Ensure get_user_requirements returns all fields from _get_requirement_select_fields
        # Fields: id, title, desc, status, priority, created_at, 
        #         assigner_name, assigner_id, assignee_name, assignee_id, 
        #         scheduled_time, comment, completed_at, attachment_path, deleted_at
        requirements_data = self.execute_with_connection(get_user_requirements, self.user_id) or []
        requirement = None
        
        for req_data in requirements_data:
            if req_data[0] == req_id: # Compare by ID
                requirement = req_data
                break
        
        if not requirement:
            messagebox.showerror("錯誤", f"找不到ID為 {req_id} 的需求單詳情。")
            return

        if len(requirement) < 14: # Should be 15 fields, attachment_path is 13, deleted_at is 14
            messagebox.showerror("錯誤", "需求單資料不完整，無法顯示詳情 (欄位數量不足)。")
            return
            
        # Unpack all 15 fields
        (req_id, title, description, status, priority, created_at, 
         assigner_name, assigner_id, assignee_name, assignee_id, 
         scheduled_time, comment, completed_at, attachment_path, deleted_at) = requirement
        
        detail_window = self.create_toplevel_window(f"需求單詳情 #{req_id}", "550x600") # Adjusted size for potential attachment
        
        # 標題
        ttk.Label(
            detail_window, 
            text=f"標題: {title}", 
            font=('Arial', 12, 'bold')
        ).pack(pady=(20, 10), padx=20, anchor=tk.W)
        
        # 發派人
        ttk.Label(
            detail_window, 
            text=f"發派人: {assigner_name}"
        ).pack(pady=5, padx=20, anchor=tk.W)
        
        # 狀態
        status_text = self.get_status_display_text(status)
        status_label = ttk.Label(
            detail_window, 
            text=f"狀態: {status_text}"
        )
        status_label.pack(pady=5, padx=20, anchor=tk.W)
        
        # 根據狀態設置顏色
        if status == 'reviewing':
            status_label.configure(foreground="blue")
        elif status == 'completed':
            status_label.configure(foreground="green")
        elif status == 'invalid':
            status_label.configure(foreground="gray")
        
        # 緊急程度
        priority_text = "緊急" if priority == "urgent" else "普通"
        priority_label = ttk.Label(
            detail_window, 
            text=f"緊急程度: {priority_text}"
        )
        priority_label.pack(pady=5, padx=20, anchor=tk.W)
        
        if priority == "urgent":
            priority_label.configure(foreground="red")
        
        # 發派時間（實際發派時間）
        if isinstance(created_at, str):
            try:
                date_obj = datetime.datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S")
                date_text = date_obj.strftime("%Y-%m-%d %H:%M")
            except ValueError:
                date_text = created_at
        else:
            date_text = created_at
            
        ttk.Label(
            detail_window, 
            text=f"實際發派時間: {date_text}",
            font=('Arial', 10)
        ).pack(pady=5, padx=20, anchor=tk.W)
        
        # 如果有預約時間，顯示預約時間
        if scheduled_time:
            if isinstance(scheduled_time, str):
                try:
                    date_obj = datetime.datetime.strptime(scheduled_time, "%Y-%m-%d %H:%M:%S")
                    scheduled_text = date_obj.strftime("%Y-%m-%d %H:%M")
                except ValueError:
                    scheduled_text = scheduled_time
            else:
                scheduled_text = scheduled_time
                
            ttk.Label(
                detail_window, 
                text=f"預約發派時間: {scheduled_text}",
                foreground="blue"
            ).pack(pady=5, padx=20, anchor=tk.W)
        
        # 內容標題
        ttk.Label(
            detail_window, 
            text="需求內容:", 
            font=('Arial', 10, 'bold')
        ).pack(pady=(15, 5), padx=20, anchor=tk.W)
        
        # 內容文本框
        content_frame = ttk.Frame(detail_window)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        
        content_text = tk.Text(content_frame, wrap=tk.WORD, height=8)
        content_text.insert(tk.END, description)
        content_text.config(state=tk.DISABLED)  # 設為只讀
        
        scrollbar = ttk.Scrollbar(content_frame, command=content_text.yview)
        content_text.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        content_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 如果有完成說明，顯示完成說明
        if comment and (status == 'reviewing' or status == 'completed'):
            ttk.Label(
                detail_window, 
                text="完成情況:", 
                font=('Arial', 10, 'bold')
            ).pack(pady=(10, 5), padx=20, anchor=tk.W)
            
            comment_frame = ttk.Frame(detail_window)
            comment_frame.pack(fill=tk.X, padx=20, pady=5)
            
            comment_text_widget = tk.Text(comment_frame, wrap=tk.WORD, height=4) # Renamed to avoid conflict
            comment_text_widget.insert(tk.END, comment)
            comment_text_widget.config(state=tk.DISABLED)  # 設為只讀
            
            comment_scroll = ttk.Scrollbar(comment_frame, command=comment_text_widget.yview)
            comment_text_widget.configure(yscrollcommand=comment_scroll.set)
            
            comment_scroll.pack(side=tk.RIGHT, fill=tk.Y)
            comment_text_widget.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 顯示附件 (New)
        if attachment_path:
            ttk.Label(detail_window, text="附件:", font=('Arial', 10, 'bold')).pack(pady=(10,0), padx=20, anchor=tk.W)
            
            attachment_display_frame = ttk.Frame(detail_window)
            attachment_display_frame.pack(fill=tk.X, padx=20, pady=(0,5))

            filename = os.path.basename(attachment_path)
            attach_label = ttk.Label(attachment_display_frame, text=filename, foreground="blue", cursor="hand2")
            attach_label.pack(side=tk.LEFT, padx=(0,10))
            # Pass the relative path from DB directly to _open_attachment
            attach_label.bind("<Button-1>", lambda e, ap=attachment_path: self._open_attachment(ap, detail_window))
        
        # 按鈕框架
        button_frame = ttk.Frame(detail_window)
        button_frame.pack(fill=tk.X, pady=15, padx=20)
        
        # 左側按鈕框架
        left_button_frame = ttk.Frame(button_frame)
        left_button_frame.pack(side=tk.LEFT, fill=tk.X, padx=20)
        
        # 如果需求單狀態是「未完成」，顯示提交按鈕
        if status == 'pending':
            ttk.Button(
                left_button_frame, 
                text="提交完成情況", 
                command=lambda: [detail_window.destroy(), self.submit_requirement()]
            ).pack(side=tk.LEFT, padx=5)
        
        # 關閉按鈕
        ttk.Button(
            button_frame, 
            text="關閉", 
            command=detail_window.destroy
        ).pack(side=tk.RIGHT, padx=20)

    def toggle_schedule_frame(self):
        """切換預約發派框架的顯示狀態"""
        if self.dispatch_method_var.get() == "scheduled":
            self.schedule_frame.grid()
        else:
            self.schedule_frame.grid_remove()

    def select_dispatch_attachment(self):
        """選擇需求單附件 (管理員發派時)"""
        filepath = filedialog.askopenfilename()
        if filepath:
            self.selected_attachment_source_path = filepath
            self.attachment_path_var.set(os.path.basename(filepath))
        else:
            self.selected_attachment_source_path = None
            self.attachment_path_var.set("")

    def create_requirement(self):
        """創建並發派需求單"""
        # 獲取選中的員工ID
        selected_staff = self.staff_combobox.get()
        if not selected_staff:
            messagebox.showerror("錯誤", "請選擇指派員工")
            return
        
        try:
            staff_id = int(selected_staff.split("ID:")[1].replace(")", ""))
        except (ValueError, IndexError):
            messagebox.showerror("錯誤", "員工選擇格式錯誤")
            return
        
        title = self.title_entry.get().strip()
        description = self.desc_text.get("1.0", tk.END).strip()
        priority = self.priority_var.get()
        
        if not title or not description:
            messagebox.showerror("錯誤", "標題和內容不能為空")
            return
            
        # 處理預約發派邏輯
        scheduled_time = None
        if self.dispatch_method_var.get() == "scheduled":
            try:
                year = int(self.year_var.get())
                month = int(self.month_var.get())
                day = int(self.day_var.get())
                hour = int(self.hour_var.get())
                minute = int(self.minute_var.get())
                
                # 檢查日期是否有效
                scheduled_datetime = datetime.datetime(year, month, day, hour, minute)
                
                # 確保預約時間在未來
                if scheduled_datetime <= datetime.datetime.now():
                    messagebox.showerror("錯誤", "預約時間必須在當前時間之後")
                    return
                    
                scheduled_time = scheduled_datetime.strftime("%Y-%m-%d %H:%M:%S")
                
            except ValueError as e:
                messagebox.showerror("錯誤", f"日期時間格式不正確: {e}")
                return

        req_id = self.execute_with_connection(
            create_requirement,
            title,
            description,
            self.user_id,
            staff_id,
            priority,
            scheduled_time,
            None  # 不再使用附件
        )

        if req_id:
            priority_text = "緊急" if priority == "urgent" else "普通"
            
            if scheduled_time:
                message = f"需求單 #{req_id} (緊急程度: {priority_text}) 已設定於 {scheduled_time} 發派"
            else:
                message = f"需求單 #{req_id} (緊急程度: {priority_text}) 已成功派發"
                
            messagebox.showinfo("成功", message)
            self.title_entry.delete(0, tk.END)
            self.desc_text.delete("1.0", tk.END)
            self.priority_var.set("normal")  # 重置為普通優先級
            self.dispatch_method_var.set("immediate")  # 重置為立即發派
            self.toggle_schedule_frame()  # 更新UI顯示
        else:
            messagebox.showerror("錯誤", "派發需求單失敗")

    def close(self):
        """關閉需求單管理界面"""
        try:
            # 關閉可能打開的詳情視窗
            for widget in self.root.winfo_children():
                if isinstance(widget, tk.Toplevel):
                    try:
                        widget.destroy()
                    except:
                        pass
                        
            # 清理管理員界面元素
            if hasattr(self, 'admin_frame') and self.admin_frame:
                try:
                    self.admin_frame.pack_forget()
                    self.admin_frame = None
                except:
                    pass
                
            if hasattr(self, 'admin_notebook') and self.admin_notebook:
                try:
                    self.admin_notebook.pack_forget()
                    self.admin_notebook = None
                except:
                    pass
                
            # 清理普通員工界面元素
            if hasattr(self, 'staff_frame') and self.staff_frame:
                try:
                    self.staff_frame.pack_forget()
                    self.staff_frame = None
                except:
                    pass
                
            # 清理追蹤的視窗
            for window in self.open_windows:
                try:
                    if window.winfo_exists():
                        window.destroy()
                except:
                    pass
            self.open_windows.clear()
                
        except Exception as e:
            print(f"關閉需求單管理界面時發生錯誤: {e}")
            # 即使發生錯誤也要確保資源被清理
            try:
                for window in self.open_windows:
                    if window.winfo_exists():
                        window.destroy()
                self.open_windows.clear()
            except:
                pass

    def load_admin_dispatched_requirements(self):
        """載入管理員已發派的需求單數據"""
        for item in self.admin_dispatched_treeview.get_children():
            self.admin_dispatched_treeview.delete(item)
            
        status_filter = self.status_filter_var.get()
        staff_filter = self.staff_filter_var.get()
        
        staff_id = None
        if staff_filter != "all" and "(" in staff_filter and ")" in staff_filter:
            try:
                # Values are like "員工名 (ID)"
                staff_id_str = staff_filter.split("(")[1].split(")")[0]
                staff_id = int(staff_id_str)
            except (ValueError, IndexError) as e:
                print(f"解析員工ID錯誤 (dispatched_tab): {e}, 原始字串: {staff_filter}")
                staff_id = None 
        
        requirements = []
        if staff_id is not None: 
            requirements = self.execute_with_connection(get_admin_requirements_by_staff, self.user_id, staff_id) or []
        else:
            requirements = self.execute_with_connection(get_admin_dispatched_requirements, self.user_id) or []
        
        for req in requirements:
            try:
                if len(req) < 15:
                    print(f"跳過不完整的管理員已發派需求單記錄 (欄位數 {len(req)}): {req}")
                    continue
                
                (req_id, title, description, status, priority, created_at, 
                 assigner_name, assigner_id, assignee_name, assignee_id,
                 scheduled_time, comment, completed_at, attachment_path, deleted_at) = req

                if status_filter != "all" and status != status_filter:
                    continue
                    
                status_text = self.get_status_display_text(status)
                priority_text = "緊急" if priority == "urgent" else "普通"
                
                created_at_display = created_at
                if isinstance(created_at, str):
                    try:
                        date_obj = datetime.datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S")
                        created_at_display = date_obj.strftime("%Y-%m-%d %H:%M")
                    except ValueError:
                        pass 
                
                item_id_val = self.admin_dispatched_treeview.insert(
                    "", tk.END, 
                    values=(req_id, title, assignee_name, status_text, priority_text, created_at_display)
                )
                
                if status == 'reviewing':
                    self.admin_dispatched_treeview.item(item_id_val, tags=('reviewing',))
                elif status == 'completed':
                    self.admin_dispatched_treeview.item(item_id_val, tags=('completed',))
                elif status == 'invalid':
                    self.admin_dispatched_treeview.item(item_id_val, tags=('invalid',))
            except Exception as e:
                print(f"載入管理員已發派需求單列表時發生錯誤: {e}, 數據: {req}")
                import traceback
                print(traceback.format_exc())
                    
        self.admin_dispatched_treeview.tag_configure('reviewing', background='#d4edff')
        self.admin_dispatched_treeview.tag_configure('completed', background='#e6ffe6')
        self.admin_dispatched_treeview.tag_configure('invalid', background='#f0f0f0')

    def load_admin_scheduled_requirements(self):
        """載入管理員預約發派的需求單數據"""
        # 清空現有數據
        for item in self.admin_scheduled_treeview.get_children():
            self.admin_scheduled_treeview.delete(item)
            
        # 獲取員工過濾條件
        staff_filter = self.scheduled_staff_filter_var.get()
        
        # 解析員工ID（如果選擇了特定員工）
        staff_id = None
        if staff_filter != "all" and "(" in staff_filter and ")" in staff_filter:
            try:
                staff_id = int(staff_filter.split("(")[1].split(")")[0])
            except (ValueError, IndexError):
                pass
            
        # 獲取數據
        if staff_id:
            # 按特定員工篩選
            requirements = self.execute_with_connection(get_admin_scheduled_by_staff, self.user_id, staff_id) or []
        else:
            # 獲取所有需求單
            requirements = self.execute_with_connection(get_admin_scheduled_requirements, self.user_id) or []
        
        # 添加數據到表格
        for req in requirements:
            try:
                # 正確解析所有15個欄位
                (req_id, title, description, status, priority, created_at, 
                 assigner_name, assigner_id, assignee_name, assignee_id,
                 scheduled_time, comment, completed_at, attachment_path, deleted_at) = req
                
                # 格式化緊急程度
                priority_text = "緊急" if priority == "urgent" else "普通"
                
                # 格式化時間
                if isinstance(scheduled_time, str):
                    try:
                        date_obj = datetime.datetime.strptime(scheduled_time, "%Y-%m-%d %H:%M:%S")
                        scheduled_text = date_obj.strftime("%Y-%m-%d %H:%M")
                    except ValueError:
                        scheduled_text = scheduled_time
                else:
                    scheduled_text = scheduled_time
                    
                # 插入數據
                self.admin_scheduled_treeview.insert(
                    "", tk.END, 
                    values=(req_id, title, assignee_name, priority_text, scheduled_text)
                )
            except Exception as e:
                print(f"載入預約發派需求單時發生錯誤: {e}, 數據: {req}")
                import traceback
                print(traceback.format_exc())

    def show_dispatched_details(self, event):
        """顯示已發派需求單詳情"""
        selected_item = self.admin_dispatched_treeview.selection()
        if not selected_item:
            return
        item = self.admin_dispatched_treeview.item(selected_item)
        req_id = item['values'][0]
        
        requirements_data = self.execute_with_connection(get_admin_dispatched_requirements, self.user_id) or []
        requirement = None
        for req_data in requirements_data:
            if req_data[0] == req_id:
                requirement = req_data
                break
        
        if not requirement:
            messagebox.showerror("錯誤", f"找不到ID為 {req_id} 的需求單詳情。")
            return

        if len(requirement) < 15:
            messagebox.showerror("錯誤", "需求單資料不完整 (欄位不足)。")
            return
        
        (req_id, title, description, status, priority, created_at, 
         assigner_name, assigner_id, assignee_name, assignee_id,
         scheduled_time, comment, completed_at, attachment_path, deleted_at) = requirement
        
        status_text = self.get_status_display_text(status)
        priority_text = "緊急" if priority == "urgent" else "普通"
    
        detail_window = self.create_toplevel_window(f"需求單詳情 #{req_id}", "600x600") # Adjusted size
        
        ttk.Label(detail_window, text=title, font=('Arial', 14, 'bold')).pack(pady=(20, 10), padx=20, anchor=tk.W)
        details_frame = ttk.Frame(detail_window)
        details_frame.pack(fill=tk.X, padx=20, pady=5)
        left_details = ttk.Frame(details_frame)
        left_details.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(left_details, text=f"狀態: {status_text}", font=('Arial', 10)).pack(pady=2, anchor=tk.W)
        ttk.Label(left_details, text=f"緊急程度: {priority_text}", font=('Arial', 10), foreground="red" if priority == "urgent" else "black").pack(pady=2, anchor=tk.W)
        ttk.Label(left_details, text=f"發派時間: {created_at}", font=('Arial', 10)).pack(pady=2, anchor=tk.W)
        right_details = ttk.Frame(details_frame)
        right_details.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        ttk.Label(right_details, text=f"指派給: {assignee_name}", font=('Arial', 10)).pack(pady=2, anchor=tk.W)
        if completed_at and status in ["reviewing", "completed"]: # Ensure completed_at is not None
            ttk.Label(right_details, text=f"完成時間: {completed_at}", font=('Arial', 10)).pack(pady=2, anchor=tk.W)
        
        ttk.Separator(detail_window, orient='horizontal').pack(fill=tk.X, padx=20, pady=10)
        ttk.Label(detail_window, text="需求內容:", font=('Arial', 10, 'bold')).pack(pady=(5, 5), padx=20, anchor=tk.W)
        content_frame = ttk.Frame(detail_window)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        content_text = tk.Text(content_frame, wrap=tk.WORD, height=6) # Adjusted height
        content_text.insert(tk.END, description)
        content_text.config(state=tk.DISABLED)
        scrollbar = ttk.Scrollbar(content_frame, command=content_text.yview)
        content_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        content_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        if comment:
            ttk.Label(detail_window, text="完成情況說明:", font=('Arial', 10, 'bold')).pack(pady=(10, 5), padx=20, anchor=tk.W)
            comment_frame = ttk.Frame(detail_window)
            comment_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5) # Allow expansion
            comment_text_widget = tk.Text(comment_frame, wrap=tk.WORD, height=3) # Adjusted height
            comment_text_widget.insert(tk.END, comment)
            comment_text_widget.config(state=tk.DISABLED)
            comment_scrollbar = ttk.Scrollbar(comment_frame, command=comment_text_widget.yview)
            comment_text_widget.configure(yscrollcommand=comment_scrollbar.set)
            comment_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            comment_text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        if attachment_path:
            ttk.Label(detail_window, text="附件:", font=('Arial', 10, 'bold')).pack(pady=(10,0), padx=20, anchor=tk.W)
            attachment_display_frame = ttk.Frame(detail_window)
            attachment_display_frame.pack(fill=tk.X, padx=20, pady=(0,5))
            filename = os.path.basename(attachment_path)
            attach_label = ttk.Label(attachment_display_frame, text=filename, foreground="blue", cursor="hand2")
            attach_label.pack(side=tk.LEFT, padx=(0,10))
            attach_label.bind("<Button-1>", lambda e, ap=attachment_path: self._open_attachment(ap, detail_window))
        
        button_frame = ttk.Frame(detail_window)
        button_frame.pack(fill=tk.X, pady=15, padx=20)
        left_button_frame = ttk.Frame(button_frame)
        left_button_frame.pack(side=tk.LEFT, fill=tk.X)
        if status == "reviewing":
            ttk.Button(left_button_frame, text="審核通過", command=lambda: self.perform_approve_requirement(req_id, detail_window)).pack(side=tk.LEFT, padx=5)
            ttk.Button(left_button_frame, text="退回修改", command=lambda: self.perform_reject_requirement(req_id, detail_window)).pack(side=tk.LEFT, padx=5)
        if status != "invalid":
            ttk.Button(left_button_frame, text="設為失效", command=lambda: self.perform_invalidate_requirement(req_id, detail_window)).pack(side=tk.LEFT, padx=5)
        ttk.Button(left_button_frame, text="刪除需求單", command=lambda: self.perform_delete_requirement(req_id, detail_window)).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="關閉", command=detail_window.destroy).pack(side=tk.RIGHT)
        
    def show_scheduled_details(self, event):
        """顯示預約發派需求單詳情"""
        selected_item = self.admin_scheduled_treeview.selection()
        if not selected_item:
            return
        item = self.admin_scheduled_treeview.item(selected_item)
        req_id = item['values'][0]
        
        requirements_data = self.execute_with_connection(get_admin_scheduled_requirements, self.user_id) or []
        requirement = None
        for req_data in requirements_data:
            if req_data[0] == req_id:
                requirement = req_data
                break
        
        if not requirement:
            messagebox.showerror("錯誤", f"找不到ID為 {req_id} 的預約需求單詳情。")
            return

        if len(requirement) < 15: # Expect 15 fields
            messagebox.showerror("錯誤", "預約需求單資料不完整 (欄位不足)。")
            return
            
        (req_id, title, description, status, priority, created_at, 
         assigner_name, assigner_id, assignee_name, assignee_id,
         scheduled_time, comment, completed_at, attachment_path, deleted_at) = requirement
        
        priority_text = "緊急" if priority == "urgent" else "普通" # Use 'priority' not item['values'][3]

        detail_window = self.create_toplevel_window(f"預約需求單詳情 #{req_id}", "600x550") # Adjusted size
        
        ttk.Label(detail_window, text=title, font=('Arial', 14, 'bold')).pack(pady=(20, 10), padx=20, anchor=tk.W)
        details_frame = ttk.Frame(detail_window)
        details_frame.pack(fill=tk.X, padx=20, pady=5)
        left_details = ttk.Frame(details_frame)
        left_details.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(left_details, text=f"緊急程度: {priority_text}", font=('Arial', 10), foreground="red" if priority == "urgent" else "black").pack(pady=2, anchor=tk.W)
        # Display scheduled_time, as this is a scheduled requirement
        if scheduled_time:
             ttk.Label(left_details, text=f"預約發派時間: {scheduled_time}", font=('Arial', 10)).pack(pady=2, anchor=tk.W)
        else: # Fallback if scheduled_time is somehow null for a scheduled item
             ttk.Label(left_details, text=f"發派時間: {created_at}", font=('Arial', 10)).pack(pady=2, anchor=tk.W)


        right_details = ttk.Frame(details_frame)
        right_details.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        ttk.Label(right_details, text=f"指派給: {assignee_name}", font=('Arial', 10)).pack(pady=2, anchor=tk.W)
        
        ttk.Separator(detail_window, orient='horizontal').pack(fill=tk.X, padx=20, pady=10)
        ttk.Label(detail_window, text="需求內容:", font=('Arial', 10, 'bold')).pack(pady=(5, 5), padx=20, anchor=tk.W)
        content_frame = ttk.Frame(detail_window)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        content_text = tk.Text(content_frame, wrap=tk.WORD, height=8)
        content_text.insert(tk.END, description)
        content_text.config(state=tk.DISABLED)
        scrollbar = ttk.Scrollbar(content_frame, command=content_text.yview)
        content_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        content_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Comment and Completed_at are unlikely for not-yet-dispatched scheduled items, but show if present
        if comment:
            ttk.Label(detail_window, text="(預約時)備註/說明:", font=('Arial', 10, 'bold')).pack(pady=(10, 5), padx=20, anchor=tk.W)
            comment_display_frame = ttk.Frame(detail_window)
            comment_display_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
            comment_display_text = tk.Text(comment_display_frame, wrap=tk.WORD, height=3)
            comment_display_text.insert(tk.END, comment)
            comment_display_text.config(state=tk.DISABLED)
            comment_display_scrollbar = ttk.Scrollbar(comment_display_frame, command=comment_display_text.yview)
            comment_display_text.configure(yscrollcommand=comment_display_scrollbar.set)
            comment_display_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            comment_display_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        if attachment_path: # From the initial unpacking
            ttk.Label(detail_window, text="附件:", font=('Arial', 10, 'bold')).pack(pady=(10,0), padx=20, anchor=tk.W)
            attachment_display_frame = ttk.Frame(detail_window)
            attachment_display_frame.pack(fill=tk.X, padx=20, pady=(0,5))
            filename = os.path.basename(attachment_path)
            attach_label = ttk.Label(attachment_display_frame, text=filename, foreground="blue", cursor="hand2")
            attach_label.pack(side=tk.LEFT, padx=(0,10))
            attach_label.bind("<Button-1>", lambda e, ap=attachment_path: self._open_attachment(ap, detail_window))
        
        # Remove the old action_button_frame and its contents from show_scheduled_details
        # It was adding approve/reject buttons which are not applicable here.
        # Only "Cancel Scheduled" and "Close" are relevant.

        button_frame = ttk.Frame(detail_window) # Re-create or use a new name for clarity
        button_frame.pack(fill=tk.X, pady=15, padx=20)
        
        ttk.Button(button_frame, text="取消預約發派", command=lambda: self.perform_cancel_scheduled(req_id, detail_window)).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="關閉", command=detail_window.destroy).pack(side=tk.RIGHT)

    def cancel_scheduled_requirement(self):
        """取消預約發派需求單"""
        # 獲取選中的項目
        selected_item = self.admin_scheduled_treeview.selection()
        if not selected_item:
            messagebox.showwarning("警告", "請先選擇要取消的預約需求單")
            return
            
        item = self.admin_scheduled_treeview.item(selected_item)
        req_id = item['values'][0]
        
        # 確認取消
        confirm = messagebox.askyesno("確認取消", "確定要取消此預約發派需求單嗎？此操作無法復原！")
        if confirm:
            self.perform_cancel_scheduled(req_id)
            
    def perform_cancel_scheduled(self, req_id, window_to_close=None):
        """執行取消預約發派"""
        if self.execute_with_connection(cancel_scheduled_requirement, req_id):
            messagebox.showinfo("成功", "已成功取消預約發派需求單")
            if window_to_close:
                window_to_close.destroy()
            self.load_admin_scheduled_requirements()
        else:
            messagebox.showerror("錯誤", "取消預約發派失敗") 

    def get_status_display_text(self, status):
        """獲取狀態的顯示文字
        
        Args:
            status: 狀態代碼
            
        Returns:
            str: 顯示文字
        """
        status_map = {
            'not_dispatched': '未發派',
            'pending': '未完成',
            'reviewing': '待審核',
            'completed': '已完成',
            'invalid': '已失效'
        }
        return status_map.get(status, status) 

    def submit_requirement(self):
        """員工提交需求單完成情況"""
        # 檢查用戶ID是否有效
        if self.user_id is None:
            messagebox.showerror("錯誤", "無法獲取用戶ID，請重新登錄")
            return
            
        # 獲取選中的項目
        selected_item = self.staff_req_treeview.selection()
        if not selected_item:
            messagebox.showwarning("警告", "請先選擇要提交的需求單")
            return
            
        item = self.staff_req_treeview.item(selected_item)
        req_id = item['values'][0]
        req_title = item['values'][1]  # 獲取需求單標題
        status = item['values'][3]  # 獲取目前狀態文字
        
        # 檢查狀態是否為「未完成」
        if status != "未完成":
            messagebox.showwarning("警告", "只能提交狀態為「未完成」的需求單")
            return
            
        # 創建提交對話框，使用需求單標題作為視窗標題
        submit_window = self.create_toplevel_window(req_title, "550x350")
        
        # 標題
        ttk.Label(
            submit_window, 
            text="請說明需求單完成情況:", 
            font=('Arial', 12, 'bold')
        ).pack(pady=(20, 10), padx=20, anchor=tk.W)
        
        # 說明文本框
        comment_frame = ttk.Frame(submit_window)
        comment_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        
        comment_text = tk.Text(comment_frame, wrap=tk.WORD, height=8)
        
        scrollbar = ttk.Scrollbar(comment_frame, command=comment_text.yview)
        comment_text.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        comment_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 按鈕框架
        button_frame = ttk.Frame(submit_window)
        button_frame.pack(fill=tk.X, pady=15, padx=20)
        
        # 取消按鈕
        ttk.Button(
            button_frame, 
            text="取消", 
            command=submit_window.destroy
        ).pack(side=tk.LEFT)
        
        # 提交按鈕
        ttk.Button(
            button_frame, 
            text="提交完成情況", 
            command=lambda: self.perform_submit_requirement(
                req_id, 
                comment_text.get("1.0", tk.END).strip(), 
                submit_window
            )
        ).pack(side=tk.RIGHT)

    def perform_submit_requirement(self, req_id, comment, window):
        """執行提交需求單操作"""
        if not comment:
            messagebox.showwarning("警告", "請填寫完成情況說明", parent=window)
            return
            
        if self.execute_with_connection(submit_requirement, req_id, comment, None):
            messagebox.showinfo("成功", "需求單已提交，等待管理員審核")
            window.destroy()
            self.load_user_requirements()  # 重新加載需求單列表
        else:
            messagebox.showerror("錯誤", "提交需求單失敗")

    def perform_approve_requirement(self, req_id, window=None):
        """執行審核通過需求單"""
        confirm = messagebox.askyesno("確認審核", "確定要審核通過此需求單嗎？")
        if confirm:
            if self.execute_with_connection(approve_requirement, req_id):
                messagebox.showinfo("成功", "需求單已審核通過，狀態已改為「已完成」")
                if window:
                    window.destroy()
                # 重新載入已發派和待審核需求單列表
                self.load_admin_dispatched_requirements()
                if hasattr(self, 'load_admin_reviewing_requirements'):
                    self.load_admin_reviewing_requirements()
            else:
                messagebox.showerror("錯誤", "審核需求單失敗")
                
    def perform_reject_requirement(self, req_id, window=None):
        """執行退回需求單"""
        confirm = messagebox.askyesno("確認退回", "確定要退回此需求單嗎？狀態將改回「未完成」")
        if confirm:
            if self.execute_with_connection(reject_requirement, req_id):
                messagebox.showinfo("成功", "需求單已退回，狀態已改為「未完成」")
                if window:
                    window.destroy()
                # 重新載入已發派和待審核需求單列表
                self.load_admin_dispatched_requirements()
                if hasattr(self, 'load_admin_reviewing_requirements'):
                    self.load_admin_reviewing_requirements()
            else:
                messagebox.showerror("錯誤", "退回需求單失敗")
                
    def perform_invalidate_requirement(self, req_id, window=None):
        """執行使需求單失效"""
        confirm = messagebox.askyesno("確認設為失效", "確定要將此需求單設為失效嗎？此操作無法撤銷！")
        if confirm:
            if self.execute_with_connection(invalidate_requirement, req_id):
                messagebox.showinfo("成功", "需求單已設為失效")
                if window:
                    window.destroy()
                self.load_admin_dispatched_requirements()
            else:
                messagebox.showerror("錯誤", "設為失效失敗") 

    def perform_delete_requirement(self, req_id, window=None):
        """執行刪除需求單操作"""
        confirm = messagebox.askyesno("確認刪除", "確定要刪除此需求單嗎？\n刪除後可在垃圾桶中查看或恢復。")
        if confirm:
            try:
                if self.execute_with_connection(delete_requirement, req_id):
                    messagebox.showinfo("成功", "需求單已移至垃圾桶")
                    if window:
                        window.destroy()
                    self.load_admin_dispatched_requirements()
                else:
                    messagebox.showerror("錯誤", "刪除需求單失敗")
            except Exception as e:
                messagebox.showerror("錯誤", f"刪除需求單時發生異常: {str(e)}")
                print(f"刪除需求單時發生異常: {str(e)}")

    def setup_trash_tab(self, parent):
        """設置垃圾桶標籤頁"""
        # 創建框架
        trash_frame = ttk.Frame(parent, padding=10)
        trash_frame.pack(fill=tk.BOTH, expand=True)
        
        # 創建工具欄
        toolbar = ttk.Frame(trash_frame)
        toolbar.pack(fill=tk.X, pady=(0, 10))
        
        # 標題標籤
        ttk.Label(
            toolbar, 
            text="已刪除的需求單", 
            font=('Arial', 12, 'bold')
        ).pack(side=tk.LEFT)
        
        # 刷新按鈕
        ttk.Button(
            toolbar, 
            text="刷新", 
            command=self.load_deleted_requirements
        ).pack(side=tk.RIGHT, padx=5)
        
        # 創建樹狀視圖用於顯示已刪除的需求單
        columns = ("ID", "標題", "緊急程度", "刪除時間", "指派給", "狀態")
        self.trash_treeview = ttk.Treeview(trash_frame, columns=columns, show="headings", selectmode="browse")
        
        # 設置列的寬度
        self.trash_treeview.column("ID", width=50, anchor=tk.CENTER)
        self.trash_treeview.column("標題", width=200)
        self.trash_treeview.column("緊急程度", width=80, anchor=tk.CENTER)
        self.trash_treeview.column("刪除時間", width=130, anchor=tk.CENTER)
        self.trash_treeview.column("指派給", width=100, anchor=tk.CENTER)
        self.trash_treeview.column("狀態", width=80, anchor=tk.CENTER)
        
        # 設置列標題
        for col in columns:
            self.trash_treeview.heading(col, text=col)
            
        # 添加滾動條
        scrollbar = ttk.Scrollbar(trash_frame, orient="vertical", command=self.trash_treeview.yview)
        self.trash_treeview.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.trash_treeview.pack(fill=tk.BOTH, expand=True)
        
        # 綁定雙擊事件
        self.trash_treeview.bind("<Double-1>", self.show_deleted_details)
        
        # 載入已刪除的需求單
        self.load_deleted_requirements()
    
    def setup_profile_tab(self, parent):
        """設置個人資料標籤頁"""
        # 創建主框架
        profile_frame = ttk.Frame(parent, padding=20)
        profile_frame.pack(fill=tk.BOTH, expand=True)
        
        # 個人資料標題
        title_label = ttk.Label(profile_frame, text="個人資料", font=('Arial', 16, 'bold'))
        title_label.pack(pady=(0, 30))
        
        # 使用者資訊區域
        info_frame = ttk.LabelFrame(profile_frame, text="使用者資訊", padding=20)
        info_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 使用者資訊標籤 - 使用網格佈局
        info_grid = ttk.Frame(info_frame)
        info_grid.pack(fill=tk.X)
        
        # 使用者名稱
        ttk.Label(info_grid, text="使用者名稱:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky=tk.W, pady=5, padx=(0, 10))
        self.profile_username_label = ttk.Label(info_grid, text="", font=('Arial', 10))
        self.profile_username_label.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # 姓名
        ttk.Label(info_grid, text="姓名:", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky=tk.W, pady=5, padx=(0, 10))
        self.profile_name_label = ttk.Label(info_grid, text="", font=('Arial', 10))
        self.profile_name_label.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # 電子郵件
        ttk.Label(info_grid, text="電子郵件:", font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky=tk.W, pady=5, padx=(0, 10))
        self.profile_email_label = ttk.Label(info_grid, text="", font=('Arial', 10))
        self.profile_email_label.grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # 角色
        ttk.Label(info_grid, text="角色:", font=('Arial', 10, 'bold')).grid(row=3, column=0, sticky=tk.W, pady=5, padx=(0, 10))
        self.profile_role_label = ttk.Label(info_grid, text="", font=('Arial', 10))
        self.profile_role_label.grid(row=3, column=1, sticky=tk.W, pady=5)
        
        # 分隔線
        separator = ttk.Separator(profile_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=20)
        
        # 操作區域
        action_frame = ttk.LabelFrame(profile_frame, text="操作", padding=20)
        action_frame.pack(fill=tk.X)
        
        # 登出按鈕
        logout_button = ttk.Button(
            action_frame, 
            text="登出系統", 
            width=15,
            command=self.perform_logout
        )
        logout_button.pack(pady=10)
        
        # 更新個人資料顯示
        self.update_profile_display()
    
    def update_profile_display(self):
        """更新個人資料顯示"""
        if hasattr(self.current_user, 'username'):
            self.profile_username_label.config(text=self.current_user.username)
        if hasattr(self.current_user, 'name'):
            self.profile_name_label.config(text=self.current_user.name)
        if hasattr(self.current_user, 'email'):
            self.profile_email_label.config(text=self.current_user.email)
        if hasattr(self.current_user, 'role'):
            role_text = "系統管理員" if self.current_user.role == 'admin' else "一般員工"
            self.profile_role_label.config(text=role_text)
    
    def perform_logout(self):
        """執行登出操作"""
        from tkinter import messagebox
        
        # 詢問用戶是否確定要登出
        confirm = messagebox.askyesno("確認登出", "您確定要登出系統嗎？")
        
        if confirm:
            try:
                # 關閉當前需求管理器的所有資源
                self.close()
                
                # 通過事件機制觸發主程式的登出，傳遞已確認的標記
                self.root.event_generate("<<LogoutConfirmed>>")
                
            except Exception as e:
                print(f"登出時發生錯誤: {e}")
                messagebox.showerror("錯誤", f"登出時發生錯誤: {e}")

    def load_deleted_requirements(self):
        """載入已刪除的需求單"""
        # 清空現有項目
        for item in self.trash_treeview.get_children():
            self.trash_treeview.delete(item)
            
        # 從數據庫獲取已刪除的需求單
        requirements = self.execute_with_connection(get_deleted_requirements, self.user_id) or []
        
        # 填充樹狀視圖
        for req in requirements:
            req_id = req[0]
            title = req[1]
            priority = "緊急" if req[4] == "urgent" else "普通"
            deleted_time = req[8][:19] if req[8] else "-"  # 截取日期時間部分
            assignee = req[6]
            status = self.get_status_display_text(req[3])
            
            # 根據緊急程度設置標籤
            tag = "urgent" if priority == "緊急" else "normal"
            
            self.trash_treeview.insert(
                "", "end", values=(req_id, title, priority, deleted_time, assignee, status), tags=(tag,)
            )
            
        # 設置標籤顏色
        self.trash_treeview.tag_configure("urgent", foreground="red")
        self.trash_treeview.tag_configure("normal", foreground="black")

    def show_deleted_details(self, event):
        """顯示已刪除需求單的詳情"""
        # 獲取選中的項目
        selected_item = self.trash_treeview.selection()
        if not selected_item:
            return
            
        item = self.trash_treeview.item(selected_item)
        req_id = item['values'][0]
        
        # 從數據庫獲取所有已刪除的需求單
        requirements = self.execute_with_connection(get_deleted_requirements, self.user_id) or []
        
        # 查找對應的需求單
        requirement = None
        for req in requirements:
            if req[0] == req_id:
                requirement = req
                break
                
        if not requirement:
            messagebox.showerror("錯誤", "找不到需求單資訊")
            return
            
        # 解析需求單資訊
        title = requirement[1]
        description = requirement[2]
        status = self.get_status_display_text(requirement[3])
        priority = "緊急" if requirement[4] == "urgent" else "普通"
        created_time = requirement[5][:19] if requirement[5] else "-"  # 截取日期時間部分
        assignee = requirement[6]
        deleted_time = requirement[8][:19] if requirement[8] else "-"  # 截取日期時間部分
        comment = requirement[9] if requirement[9] else ""
        
        # 創建詳情視窗
        detail_window = self.create_toplevel_window(f"已刪除需求單詳情 #{req_id}", "600x500")
        
        # 標題
        ttk.Label(
            detail_window, 
            text=title, 
            font=('Arial', 14, 'bold')
        ).pack(pady=(20, 10), padx=20, anchor=tk.W)
        
        # 詳情框架
        details_frame = ttk.Frame(detail_window)
        details_frame.pack(fill=tk.X, padx=20, pady=5)
        
        # 左側詳情
        left_details = ttk.Frame(details_frame)
        left_details.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Label(
            left_details, 
            text=f"狀態: {status}", 
            font=('Arial', 10)
        ).pack(pady=2, anchor=tk.W)
        
        ttk.Label(
            left_details, 
            text=f"緊急程度: {priority}", 
            font=('Arial', 10),
            foreground="red" if priority == "緊急" else "black"
        ).pack(pady=2, anchor=tk.W)
        
        ttk.Label(
            left_details, 
            text=f"刪除時間: {deleted_time}", 
            font=('Arial', 10)
        ).pack(pady=2, anchor=tk.W)
        
        # 右側詳情
        right_details = ttk.Frame(details_frame)
        right_details.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        
        ttk.Label(
            right_details, 
            text=f"指派給: {assignee}", 
            font=('Arial', 10)
        ).pack(pady=2, anchor=tk.W)
        
        # 分隔線
        ttk.Separator(detail_window, orient='horizontal').pack(fill=tk.X, padx=20, pady=10)
        
        # 內容標題
        ttk.Label(
            detail_window, 
            text="需求內容:", 
            font=('Arial', 10, 'bold')
        ).pack(pady=(5, 5), padx=20, anchor=tk.W)
        
        # 內容文本框
        content_frame = ttk.Frame(detail_window)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        
        content_text = tk.Text(content_frame, wrap=tk.WORD, height=8)
        content_text.insert(tk.END, description)
        content_text.config(state=tk.DISABLED)  # 設為只讀
        
        scrollbar = ttk.Scrollbar(content_frame, command=content_text.yview)
        content_text.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        content_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 如果有完成情況說明，則顯示
        if comment:
            # 說明標題
            ttk.Label(
                detail_window, 
                text="完成情況說明:", 
                font=('Arial', 10, 'bold')
            ).pack(pady=(10, 5), padx=20, anchor=tk.W)
            
            # 說明文本框
            comment_frame = ttk.Frame(detail_window)
            comment_frame.pack(fill=tk.X, padx=20, pady=5)
            
            comment_text = tk.Text(comment_frame, wrap=tk.WORD, height=4)
            comment_text.insert(tk.END, comment)
            comment_text.config(state=tk.DISABLED)  # 設為只讀
            
            comment_scroll = ttk.Scrollbar(comment_frame, command=comment_text.yview)
            comment_text.configure(yscrollcommand=comment_scroll.set)
            
            comment_scroll.pack(side=tk.RIGHT, fill=tk.Y)
            comment_text.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 按鈕框架
        button_frame = ttk.Frame(detail_window)
        button_frame.pack(fill=tk.X, pady=15, padx=20)
        
        # 左側按鈕 - 恢復需求單
        ttk.Button(
            button_frame, 
            text="恢復需求單", 
            command=lambda: self.perform_restore_requirement(req_id, detail_window)
        ).pack(side=tk.LEFT, padx=5)
        
        # 右側按鈕 - 關閉按鈕
        ttk.Button(
            button_frame, 
            text="關閉", 
            command=detail_window.destroy
        ).pack(side=tk.RIGHT)

    def perform_restore_requirement(self, req_id, window=None):
        """執行恢復需求單操作"""
        confirm = messagebox.askyesno("確認恢復", "確定要恢復此需求單嗎？")
        if confirm:
            if self.execute_with_connection(restore_requirement, req_id):
                messagebox.showinfo("成功", "需求單已恢復")
                if window:
                    window.destroy()
                self.load_deleted_requirements()
                self.load_admin_dispatched_requirements()
            else:
                messagebox.showerror("錯誤", "恢復需求單失敗")

    def refresh_staff_list(self):
        """刷新員工列表"""
        try:
            # 獲取當前選中的值（如果有）
            current_selection = self.staff_var.get()
            
            # 重新從數據庫獲取員工列表
            staffs = self.execute_with_connection(get_all_staff)
            if staffs is None:
                messagebox.showerror("錯誤", "無法獲取員工列表")
                return
            
            # 更新下拉選單的值
            self.staff_combobox['values'] = [f"{staff[1]} (ID:{staff[0]})" for staff in staffs]
            
            # 如果之前有選中值，嘗試保持它
            if current_selection:
                self.staff_combobox.set(current_selection)
                
            messagebox.showinfo("成功", "員工列表已刷新")
        except Exception as e:
            messagebox.showerror("錯誤", f"刷新員工列表時發生錯誤: {str(e)}")
            print(f"刷新員工列表錯誤: {str(e)}")

    def load_admin_reviewing_requirements(self):
        """載入管理員待審核的需求單數據"""
        # 清空現有數據
        for item in self.admin_reviewing_treeview.get_children():
            self.admin_reviewing_treeview.delete(item)
            
        # 獲取所有已發派的需求單 (這些應包含完整的15個欄位)
        requirements = self.execute_with_connection(get_admin_dispatched_requirements, self.user_id) or []
        
        # 篩選狀態為「待審核」的需求單
        reviewing_requirements = [req for req in requirements if len(req) == 15 and req[3] == 'reviewing']
        
        # 添加數據到表格
        for req in reviewing_requirements:
            try:
                # 解包完整的15個欄位
                (req_id, title, description, status, priority, created_at, 
                 assigner_name, assigner_id, assignee_name, assignee_id,
                 scheduled_time, comment, completed_at, attachment_path, deleted_at) = req
                
                # 格式化緊急程度
                priority_text = "緊急" if priority == "urgent" else "普通"
                
                # 格式化時間
                date_text = created_at # Default
                if isinstance(created_at, str):
                    try:
                        date_obj = datetime.datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S")
                        date_text = date_obj.strftime("%Y-%m-%d %H:%M")
                    except ValueError:
                        pass # date_text remains original created_at
                
                # Treeview columns: ("id", "title", "assignee", "priority", "created_at")
                item_id_val = self.admin_reviewing_treeview.insert(
                    "", tk.END, 
                    values=(req_id, title, assignee_name, priority_text, date_text)
                )
                
                # 根據優先級設置行顏色
                if priority == 'urgent':
                    self.admin_reviewing_treeview.item(item_id_val, tags=('urgent',))
                
            except Exception as e:
                print(f"處理待審核需求單時發生錯誤: {e}, 數據: {req}")
                import traceback
                print(traceback.format_exc())
                    
        # 設置標籤顏色
        self.admin_reviewing_treeview.tag_configure('urgent', background='#d4edff')

    def show_reviewing_requirement_details(self, event):
        """顯示待審核需求單詳情"""
        # 獲取選中的項目
        selected_item = self.admin_reviewing_treeview.selection()
        if not selected_item:
            return
            
        item = self.admin_reviewing_treeview.item(selected_item)
        req_id_from_tree = item['values'][0] # Renamed to avoid confusion with req_id from full data
        
        # 獲取需求單詳情 (應包含完整的15個欄位)
        requirements_data = self.execute_with_connection(get_admin_dispatched_requirements, self.user_id) or []
        requirement_full = None # Renamed
        
        for req_data_full in requirements_data: # Renamed
            if len(req_data_full) == 15 and req_data_full[0] == req_id_from_tree: # Check length before access
                requirement_full = req_data_full
                break
        
        if not requirement_full:
            messagebox.showerror("錯誤", f"找不到ID為 {req_id_from_tree} 的完整需求單詳情。")
            return
            
        # 解包完整的15個欄位
        (req_id, title, description, status, priority, created_at, 
         assigner_name, assigner_id, assignee_name, assignee_id,
         scheduled_time, comment, completed_at, attachment_path, deleted_at) = requirement_full
        
        # 創建詳情對話框
        detail_window = self.create_toplevel_window(f"待審核需求單詳情 #{req_id}", "600x650")
        
        # 標題
        ttk.Label(
            detail_window, 
            text=f"標題: {title}", 
            font=('Arial', 12, 'bold')
        ).pack(pady=(20, 10), padx=20, anchor=tk.W)
        
        # 狀態標籤 (Status is 'reviewing' by definition of this tab)
        status_frame = ttk.Frame(detail_window)
        status_frame.pack(fill=tk.X, padx=20, pady=5)
        
        status_label = ttk.Label(
            status_frame, 
            text=f"狀態: {self.get_status_display_text(status)}", # Use actual status from data
            font=('Arial', 10, 'bold'),
            foreground="blue"
        )
        status_label.pack(side=tk.LEFT)
        
        # 詳情區域 - 使用網格佈局
        details_frame = ttk.Frame(detail_window)
        details_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # 左側詳情
        left_details = ttk.Frame(details_frame)
        left_details.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 緊急程度
        priority_text = "緊急" if priority == "urgent" else "普通"
        priority_label = ttk.Label(
            left_details, 
            text=f"緊急程度: {priority_text}",
            font=('Arial', 10),
        )
        priority_label.pack(pady=2, anchor=tk.W)
        
        if priority == "urgent":
            priority_label.configure(foreground="red")
        
        # 發派時間
        created_at_display = created_at
        if isinstance(created_at, str):
            try:
                created_at_display = datetime.datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S").strftime("%Y-%m-%d %H:%M")
            except ValueError: pass
        ttk.Label(
            left_details, 
            text=f"發派時間: {created_at_display}", 
            font=('Arial', 10)
        ).pack(pady=2, anchor=tk.W)
        
        # 右側詳情
        right_details = ttk.Frame(details_frame)
        right_details.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        
        # 接收人
        ttk.Label(
            right_details, 
            text=f"接收人: {assignee_name}", 
            font=('Arial', 10)
        ).pack(pady=2, anchor=tk.W)
        
        # 提交時間 (completed_at for reviewing items)
        completed_at_display = completed_at
        if isinstance(completed_at, str):
            try:
                completed_at_display = datetime.datetime.strptime(completed_at, "%Y-%m-%d %H:%M:%S").strftime("%Y-%m-%d %H:%M")
            except ValueError: pass

        ttk.Label(
            right_details, 
            text=f"提交時間: {completed_at_display if completed_at else '-'}", # Show '-' if None
            font=('Arial', 10)
        ).pack(pady=2, anchor=tk.W)
        
        # 分隔線
        ttk.Separator(detail_window, orient='horizontal').pack(fill=tk.X, padx=20, pady=10)
        
        # 內容標題
        ttk.Label(
            detail_window, 
            text="需求內容:", 
            font=('Arial', 10, 'bold')
        ).pack(pady=(5, 5), padx=20, anchor=tk.W)
        
        # 內容文本框
        content_frame = ttk.Frame(detail_window)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        
        content_text = tk.Text(content_frame, wrap=tk.WORD, height=8)
        content_text.insert(tk.END, description)
        content_text.config(state=tk.DISABLED)
        
        scrollbar = ttk.Scrollbar(content_frame, command=content_text.yview)
        content_text.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        content_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 員工完成情況說明
        if comment:
            ttk.Label(
                detail_window, 
                text="員工完成情況說明:", 
                font=('Arial', 10, 'bold')
            ).pack(pady=(10, 5), padx=20, anchor=tk.W)
            
            comment_display_frame = ttk.Frame(detail_window) # Renamed to avoid conflict with the Text widget
            comment_display_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
            
            comment_text_widget = tk.Text(comment_display_frame, wrap=tk.WORD, height=6) # Renamed
            comment_text_widget.insert(tk.END, comment)
            comment_text_widget.config(state=tk.DISABLED)
            
            comment_scrollbar = ttk.Scrollbar(comment_display_frame, command=comment_text_widget.yview) # Use renamed widget
            comment_text_widget.configure(yscrollcommand=comment_scrollbar.set) # Use renamed widget
            
            comment_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            comment_text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True) # Use renamed widget

        # 附件顯示
        if attachment_path:
            ttk.Label(detail_window, text="附件:", font=('Arial', 10, 'bold')).pack(pady=(10,0), padx=20, anchor=tk.W)
            attachment_display_frame_details = ttk.Frame(detail_window) # Unique name
            attachment_display_frame_details.pack(fill=tk.X, padx=20, pady=(0,5))

            filename = os.path.basename(attachment_path)
            attach_label = ttk.Label(attachment_display_frame_details, text=filename, foreground="blue", cursor="hand2")
            attach_label.pack(side=tk.LEFT, padx=(0,10))
            attach_label.bind("<Button-1>", lambda e, ap=attachment_path: self._open_attachment(ap, detail_window))
        
        # 按鈕框架
        button_frame = ttk.Frame(detail_window)
        button_frame.pack(fill=tk.X, pady=15, padx=20)
        
        left_button_frame = ttk.Frame(button_frame)
        left_button_frame.pack(side=tk.LEFT)
        
        ttk.Button(
            left_button_frame, 
            text="審核通過", 
            command=lambda: self.perform_approve_requirement(req_id, detail_window)
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            left_button_frame, 
            text="退回修改", 
            command=lambda: self.perform_reject_requirement(req_id, detail_window)
        ).pack(side=tk.LEFT, padx=5)
        
        right_button_frame = ttk.Frame(button_frame) # No need for this extra frame
        right_button_frame.pack(side=tk.RIGHT)
        
        ttk.Button(
            right_button_frame,  # Can be button_frame directly
            text="關閉", 
            command=detail_window.destroy
        ).pack(side=tk.RIGHT, padx=5) # This was also .pack(side=tk.RIGHT)

    def _open_attachment(self, attachment_path, window):
        try:
            if hasattr(os, 'startfile'):
                os.startfile(attachment_path)
            else:
                messagebox.showerror("錯誤", "無法打開附件")
        except Exception as e:
            messagebox.showerror("錯誤", f"打開附件時發生錯誤: {e}")
