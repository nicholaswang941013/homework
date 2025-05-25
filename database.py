import sqlite3
from sqlite3 import Error
import datetime

def create_connection():
    """建立資料庫連接"""
    conn = None
    try:
        conn = sqlite3.connect('users.db')
        return conn
    except Error as e:
        print(e)
    return conn

def get_user_by_username(conn, username):
    """根據使用者名稱獲取使用者資料"""
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE username=?", (username,))
    return cursor.fetchone()

def create_tables(conn):
    """建立所有需要的表格"""
    try:
        # 使用者表格
        conn.execute('''CREATE TABLE IF NOT EXISTS users (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        username TEXT UNIQUE NOT NULL,
                        password TEXT NOT NULL,
                        name TEXT NOT NULL,
                        email TEXT NOT NULL,
                        role TEXT NOT NULL DEFAULT 'staff'
                    );''')

        # 需求單表格 - 檢查是否需要重建表格
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(requirements)")
        columns = cursor.fetchall()
        columns_dict = {col[1]: col for col in columns}
        
        # 如果缺少必要欄位，則嘗試添加它們
        if 'comment' not in columns_dict or 'completed_at' not in columns_dict or 'is_deleted' not in columns_dict:
            try:
                if 'comment' not in columns_dict:
                    conn.execute("ALTER TABLE requirements ADD COLUMN comment TEXT")
                if 'completed_at' not in columns_dict:
                    conn.execute("ALTER TABLE requirements ADD COLUMN completed_at TIMESTAMP")
                if 'is_deleted' not in columns_dict:
                    conn.execute("ALTER TABLE requirements ADD COLUMN is_deleted INTEGER DEFAULT 0")
                conn.commit()
                print("需求單表格結構已更新")
            except Error as e:
                print(f"添加欄位時發生錯誤: {e}")
        
        # 確保需求單表格存在
        conn.execute('''CREATE TABLE IF NOT EXISTS requirements (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        title TEXT NOT NULL,
                        description TEXT NOT NULL,
                        assigner_id INTEGER NOT NULL,
                        assignee_id INTEGER NOT NULL,
                        status TEXT NOT NULL DEFAULT 'pending',
                        priority TEXT NOT NULL DEFAULT 'normal',
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        scheduled_time TIMESTAMP,
                        is_dispatched INTEGER DEFAULT 1,
                        completed_at TIMESTAMP,
                        comment TEXT,
                        is_deleted INTEGER DEFAULT 0,
                        deleted_at TIMESTAMP,
                        FOREIGN KEY (assigner_id) REFERENCES users (id),
                        FOREIGN KEY (assignee_id) REFERENCES users (id)
                    );''')
    except Error as e:
        print(e)

def initialize_database():
    """初始化資料庫"""
    conn = create_connection()
    if conn is not None:
        create_tables(conn)

        # 定義預設使用者
        default_users = [
            ('nicholas', 'nicholas941013', '王爺', 'yuxiangwang57@gmail.com', 'admin'),
            ('user1', 'user123', '張三', 'user1@example.com', 'staff'),
            ('staff1', 'staff123', '李四', 'staff1@example.com', 'staff'),
            ('staff2', 'staff123', '王五', 'staff2@example.com', 'staff')
        ]
        
        # 檢查每個使用者是否已存在，不存在則添加
        cursor = conn.cursor()
        for user in default_users:
            username = user[0]
            cursor.execute("SELECT COUNT(*) FROM users WHERE username=?", (username,))
            if cursor.fetchone()[0] == 0:
                cursor.execute(
                    "INSERT INTO users (username, password, name, email, role) VALUES (?, ?, ?, ?, ?)",
                    user
                )
                print(f"已添加預設使用者: {username}")
        
        conn.commit()
        conn.close()

def add_user(username, password, name, email, role='staff'):
    """添加新使用者
    
    Args:
        username: 使用者名稱
        password: 密碼
        name: 真實姓名
        email: 電子郵件
        role: 角色 ('admin'或'staff')
        
    Returns:
        bool: 操作是否成功
    """
    try:
        conn = create_connection()
        if conn is None:
            return False
            
        cursor = conn.cursor()
        
        # 先檢查使用者名稱是否已存在
        cursor.execute("SELECT COUNT(*) FROM users WHERE username = ?", (username,))
        if cursor.fetchone()[0] > 0:
            print(f"使用者名稱 '{username}' 已存在")
            conn.close()
            return False
            
        cursor.execute(
            "INSERT INTO users (username, password, name, email, role) VALUES (?, ?, ?, ?, ?)",
            (username, password, name, email, role)
        )
        conn.commit()
        user_id = cursor.lastrowid
        conn.close()
        
        print(f"成功添加使用者 '{username}' (ID: {user_id})")
        return True
    except Error as e:
        print(f"添加使用者時發生錯誤: {e}")
        return False

def create_requirement(conn, title, description, assigner_id, assignee_id, priority='normal', scheduled_time=None):
    """建立新的需求單
    
    Args:
        conn: 數據庫連接
        title: 需求單標題
        description: 需求單內容
        assigner_id: 指派者ID
        assignee_id: 接收者ID
        priority: 優先級 ('normal'或'urgent')
        scheduled_time: 預約發派時間，None表示立即發派
    
    Returns:
        int: 新建需求單ID
    """
    try:
        cursor = conn.cursor()
        
        # 判斷是否是預約發派
        is_dispatched = 0 if scheduled_time else 1
        
        # 設置狀態：未發派(not_dispatched)或未完成(pending)
        status = 'not_dispatched' if scheduled_time else 'pending'
        
        cursor.execute(
            """INSERT INTO requirements 
               (title, description, assigner_id, assignee_id, priority, scheduled_time, is_dispatched, status) 
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (title, description, assigner_id, assignee_id, priority, scheduled_time, is_dispatched, status)
        )
        conn.commit()
        return cursor.lastrowid
    except Error as e:
        print(e)
        return None

def get_all_staff(conn):
    """獲取所有員工列表 (排除管理員)"""
    cursor = conn.cursor()
    cursor.execute("SELECT id, name FROM users WHERE role = 'staff'")
    return cursor.fetchall()

def get_user_requirements(conn, user_id):
    """獲取指定用戶收到的需求單 (只顯示已發派的)"""
    try:
        cursor = conn.cursor()
        cursor.execute('''
            SELECT r.id, r.title, r.description, r.status, r.priority, r.created_at, u.name, r.scheduled_time, r.comment
            FROM requirements r
            JOIN users u ON r.assigner_id = u.id
            WHERE r.assignee_id = ? AND r.is_dispatched = 1 AND r.is_deleted = 0
            ORDER BY r.created_at DESC
        ''', (user_id,))
        return cursor.fetchall()
    except Error as e:
        print(e)
        return []

def get_admin_dispatched_requirements(conn, admin_id):
    """獲取管理員已發派的需求單"""
    try:
        cursor = conn.cursor()
        cursor.execute('''
            SELECT r.id, r.title, r.description, r.status, r.priority, r.created_at, 
                   u.name as assignee_name, u.id as assignee_id, r.scheduled_time, r.comment, r.completed_at
            FROM requirements r
            JOIN users u ON r.assignee_id = u.id
            WHERE r.assigner_id = ? AND r.is_dispatched = 1 AND r.is_deleted = 0
            ORDER BY r.created_at DESC
        ''', (admin_id,))
        return cursor.fetchall()
    except Error as e:
        print(e)
        return []

def get_admin_requirements_by_staff(conn, admin_id, staff_id):
    """獲取管理員發派給特定員工的需求單
    
    Args:
        conn: 數據庫連接
        admin_id: 管理員ID
        staff_id: 員工ID
        
    Returns:
        list: 需求單列表
    """
    try:
        cursor = conn.cursor()
        cursor.execute('''
            SELECT r.id, r.title, r.description, r.status, r.priority, r.created_at, 
                   u.name as assignee_name, u.id as assignee_id, r.scheduled_time, r.comment, r.completed_at
            FROM requirements r
            JOIN users u ON r.assignee_id = u.id
            WHERE r.assigner_id = ? AND r.assignee_id = ? AND r.is_dispatched = 1 AND r.is_deleted = 0
            ORDER BY r.created_at DESC
        ''', (admin_id, staff_id))
        return cursor.fetchall()
    except Error as e:
        print(e)
        return []

def get_admin_scheduled_requirements(conn, admin_id):
    """獲取管理員預約發派的需求單（未發派）"""
    try:
        cursor = conn.cursor()
        cursor.execute('''
            SELECT r.id, r.title, r.description, r.priority, r.scheduled_time, 
                   u.name as assignee_name, u.id as assignee_id
            FROM requirements r
            JOIN users u ON r.assignee_id = u.id
            WHERE r.assigner_id = ? AND r.is_dispatched = 0
            ORDER BY r.scheduled_time ASC
        ''', (admin_id,))
        return cursor.fetchall()
    except Error as e:
        print(e)
        return []

def get_admin_scheduled_by_staff(conn, admin_id, staff_id):
    """獲取管理員發派給特定員工的預約需求單（未發派）
    
    Args:
        conn: 數據庫連接
        admin_id: 管理員ID
        staff_id: 員工ID
        
    Returns:
        list: 需求單列表
    """
    try:
        cursor = conn.cursor()
        cursor.execute('''
            SELECT r.id, r.title, r.description, r.priority, r.scheduled_time, 
                   u.name as assignee_name, u.id as assignee_id
            FROM requirements r
            JOIN users u ON r.assignee_id = u.id
            WHERE r.assigner_id = ? AND r.assignee_id = ? AND r.is_dispatched = 0
            ORDER BY r.scheduled_time ASC
        ''', (admin_id, staff_id))
        return cursor.fetchall()
    except Error as e:
        print(e)
        return []

def dispatch_scheduled_requirements(conn):
    """檢查並發派到期的預約需求單"""
    try:
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cursor = conn.cursor()
        
        # 獲取所有應該發派的需求單
        cursor.execute('''
            SELECT id FROM requirements 
            WHERE is_dispatched = 0 AND scheduled_time <= ?
        ''', (current_time,))
        
        req_ids = [row[0] for row in cursor.fetchall()]
        
        # 更新狀態為已發派，並將created_at設為當前時間
        if req_ids:
            for req_id in req_ids:
                cursor.execute('''
                    UPDATE requirements 
                    SET is_dispatched = 1, created_at = ?, status = 'pending'
                    WHERE id = ?
                ''', (current_time, req_id))
            
            conn.commit()
            return len(req_ids)
        
        return 0
    except Error as e:
        print(e)
        return 0

def cancel_scheduled_requirement(conn, req_id):
    """取消預約發派的需求單"""
    try:
        cursor = conn.cursor()
        cursor.execute('''
            DELETE FROM requirements 
            WHERE id = ? AND is_dispatched = 0
        ''', (req_id,))
        conn.commit()
        return cursor.rowcount > 0
    except Error as e:
        print(e)
        return False

def submit_requirement(conn, req_id, comment):
    """員工提交需求單完成情況
    
    Args:
        conn: 數據庫連接
        req_id: 需求單ID
        comment: 員工完成情況說明
        
    Returns:
        bool: 操作是否成功
    """
    try:
        cursor = conn.cursor()
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        cursor.execute('''
            UPDATE requirements
            SET status = 'reviewing', comment = ?, completed_at = ?
            WHERE id = ? AND status = 'pending'
        ''', (comment, current_time, req_id))
        
        conn.commit()
        return cursor.rowcount > 0
    except Error as e:
        print(e)
        return False

def approve_requirement(conn, req_id):
    """管理員審核通過需求單
    
    Args:
        conn: 數據庫連接
        req_id: 需求單ID
        
    Returns:
        bool: 操作是否成功
    """
    try:
        cursor = conn.cursor()
        
        cursor.execute('''
            UPDATE requirements
            SET status = 'completed'
            WHERE id = ? AND status = 'reviewing'
        ''', (req_id,))
        
        conn.commit()
        return cursor.rowcount > 0
    except Error as e:
        print(e)
        return False

def reject_requirement(conn, req_id):
    """管理員拒絕需求單，將狀態改回未完成
    
    Args:
        conn: 數據庫連接
        req_id: 需求單ID
        
    Returns:
        bool: 操作是否成功
    """
    try:
        cursor = conn.cursor()
        
        cursor.execute('''
            UPDATE requirements
            SET status = 'pending', comment = NULL, completed_at = NULL
            WHERE id = ? AND status = 'reviewing'
        ''', (req_id,))
        
        conn.commit()
        return cursor.rowcount > 0
    except Error as e:
        print(e)
        return False

def invalidate_requirement(conn, req_id):
    """使需求單失效
    
    Args:
        conn: 數據庫連接
        req_id: 需求單ID
        
    Returns:
        bool: 操作是否成功
    """
    try:
        cursor = conn.cursor()
        
        cursor.execute('''
            UPDATE requirements
            SET status = 'invalid'
            WHERE id = ?
        ''', (req_id,))
        
        conn.commit()
        return cursor.rowcount > 0
    except Error as e:
        print(e)
        return False

def clear_all_requirements():
    """清空所有需求單
    
    Returns:
        bool: 操作是否成功
    """
    try:
        conn = create_connection()
        if conn is None:
            print("無法連接到資料庫")
            return False
            
        cursor = conn.cursor()
        
        # 獲取需求單數量
        cursor.execute("SELECT COUNT(*) FROM requirements")
        count = cursor.fetchone()[0]
        
        # 刪除所有需求單
        cursor.execute("DELETE FROM requirements")
        conn.commit()
        
        print(f"已清空資料庫中的 {count} 個需求單")
        conn.close()
        return True
    except Error as e:
        print(f"清空需求單時發生錯誤: {e}")
        return False

def delete_requirement(conn, req_id):
    """刪除需求單（移到垃圾桶）
    
    Args:
        conn: 數據庫連接
        req_id: 需求單ID
        
    Returns:
        bool: 操作是否成功
    """
    try:
        cursor = conn.cursor()
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        cursor.execute('''
            UPDATE requirements
            SET is_deleted = 1, deleted_at = ?
            WHERE id = ? AND is_deleted = 0
        ''', (current_time, req_id))
        
        conn.commit()
        return cursor.rowcount > 0
    except Error as e:
        print(e)
        return False

def restore_requirement(conn, req_id):
    """恢復已刪除的需求單
    
    Args:
        conn: 數據庫連接
        req_id: 需求單ID
        
    Returns:
        bool: 操作是否成功
    """
    try:
        cursor = conn.cursor()
        
        cursor.execute('''
            UPDATE requirements
            SET is_deleted = 0, deleted_at = NULL
            WHERE id = ? AND is_deleted = 1
        ''', (req_id,))
        
        conn.commit()
        return cursor.rowcount > 0
    except Error as e:
        print(e)
        return False

def get_deleted_requirements(conn, admin_id):
    """獲取管理員已刪除的需求單
    
    Args:
        conn: 數據庫連接
        admin_id: 管理員ID
        
    Returns:
        list: 刪除的需求單列表
    """
    try:
        cursor = conn.cursor()
        cursor.execute('''
            SELECT r.id, r.title, r.description, r.status, r.priority, r.created_at, 
                   u.name as assignee_name, u.id as assignee_id, r.deleted_at, r.comment
            FROM requirements r
            JOIN users u ON r.assignee_id = u.id
            WHERE r.assigner_id = ? AND r.is_deleted = 1
            ORDER BY r.deleted_at DESC
        ''', (admin_id,))
        return cursor.fetchall()
    except Error as e:
        print(e)
        return [] 