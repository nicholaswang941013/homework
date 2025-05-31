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
        print(f"建立資料庫連接時發生錯誤: {e}")
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

        # 需求單表格
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
                        is_dispatched INTEGER DEFAULT 1, -- 0 for scheduled, 1 for dispatched
                        completed_at TIMESTAMP,
                        comment TEXT,
                        attachment_path TEXT,          -- New field for attachment
                        is_deleted INTEGER DEFAULT 0,    -- 0 for not deleted, 1 for deleted
                        deleted_at TIMESTAMP,
                        FOREIGN KEY (assigner_id) REFERENCES users (id),
                        FOREIGN KEY (assignee_id) REFERENCES users (id)
                    );''')
        
        # 檢查並添加欄位 (確保兼容舊資料庫)
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(requirements)")
        columns = {col[1]: col for col in cursor.fetchall()}

        if 'comment' not in columns:
            conn.execute("ALTER TABLE requirements ADD COLUMN comment TEXT")
            print("已為 requirements 表添加 comment 欄位")
        if 'completed_at' not in columns:
            conn.execute("ALTER TABLE requirements ADD COLUMN completed_at TIMESTAMP")
            print("已為 requirements 表添加 completed_at 欄位")
        if 'is_deleted' not in columns:
            conn.execute("ALTER TABLE requirements ADD COLUMN is_deleted INTEGER DEFAULT 0")
            print("已為 requirements 表添加 is_deleted 欄位")
        if 'deleted_at' not in columns:
            conn.execute("ALTER TABLE requirements ADD COLUMN deleted_at TIMESTAMP")
            print("已為 requirements 表添加 deleted_at 欄位")
        if 'attachment_path' not in columns:
            conn.execute("ALTER TABLE requirements ADD COLUMN attachment_path TEXT")
            print("已為 requirements 表添加 attachment_path 欄位")
            
        conn.commit()

    except Error as e:
        print(f"建立或更新表格時發生錯誤: {e}")

def initialize_database():
    """初始化資料庫"""
    conn = create_connection()
    if conn is not None:
        try:
            create_tables(conn)

            default_users = [
                ('nicholas', 'nicholas941013', '王爺', 'yuxiangwang57@gmail.com', 'admin'),
                ('user1', 'user123', '張三', 'user1@example.com', 'staff'),
                ('staff1', 'staff123', '李四', 'staff1@example.com', 'staff'),
                ('staff2', 'staff123', '王五', 'staff2@example.com', 'staff')
            ]
            
            cursor = conn.cursor()
            for user_data in default_users:
                cursor.execute("SELECT COUNT(*) FROM users WHERE username=?", (user_data[0],))
                if cursor.fetchone()[0] == 0:
                    cursor.execute(
                        "INSERT INTO users (username, password, name, email, role) VALUES (?, ?, ?, ?, ?)",
                        user_data
                    )
                    print(f"已添加預設使用者: {user_data[0]}")
            
            conn.commit()
        except Error as e:
            print(f"初始化資料庫時發生錯誤: {e}")
        finally:
            conn.close()

def add_user(username, password, name, email, role='staff'):
    """添加新使用者"""
    conn = create_connection()
    if not conn: 
        return False
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM users WHERE username = ?", (username,))
        if cursor.fetchone()[0] > 0:
            print(f"使用者名稱 '{username}' 已存在")
            return False
            
        cursor.execute(
            "INSERT INTO users (username, password, name, email, role) VALUES (?, ?, ?, ?, ?)",
            (username, password, name, email, role)
        )
        conn.commit()
        print(f"成功添加使用者 '{username}' (ID: {cursor.lastrowid})")
        return True
    except Error as e:
        print(f"添加使用者時發生錯誤: {e}")
        return False
    finally:
        if conn: 
            conn.close()

def create_requirement(conn, title, description, assigner_id, assignee_id, priority='normal', scheduled_time=None, attachment_path=None):
    """建立新的需求單"""
    try:
        cursor = conn.cursor()
        is_dispatched = 0 if scheduled_time else 1
        status = 'not_dispatched' if scheduled_time else 'pending'
        
        cursor.execute(
            """INSERT INTO requirements 
               (title, description, assigner_id, assignee_id, priority, scheduled_time, is_dispatched, status, attachment_path) 
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (title, description, assigner_id, assignee_id, priority, scheduled_time, is_dispatched, status, attachment_path)
        )
        conn.commit()
        return cursor.lastrowid
    except Error as e:
        print(f"建立需求單時發生錯誤: {e}")
        return None

def get_all_staff(conn):
    """獲取所有員工列表 (排除管理員)"""
    cursor = conn.cursor()
    cursor.execute("SELECT id, name FROM users WHERE role = 'staff'")
    return cursor.fetchall()

# Helper to construct SELECT query for requirements for consistency
def _get_requirement_select_fields():
    return """
        r.id, r.title, r.description, r.status, r.priority, r.created_at, 
        assigner_user.name as assigner_name, assigner_user.id as assigner_id, 
        assignee_user.name as assignee_name, assignee_user.id as assignee_id,
        r.scheduled_time, r.comment, r.completed_at, r.attachment_path, r.deleted_at
    """

def _get_requirement_joins():
    return """
        JOIN users assigner_user ON r.assigner_id = assigner_user.id
        JOIN users assignee_user ON r.assignee_id = assignee_user.id
    """

def get_user_requirements(conn, user_id):
    """獲取指定用戶收到的需求單 (只顯示已發派且未刪除的)"""
    try:
        cursor = conn.cursor()
        sql = f'''
            SELECT {_get_requirement_select_fields()}
            FROM requirements r
            {_get_requirement_joins()}
            WHERE r.assignee_id = ? AND r.is_dispatched = 1 AND r.is_deleted = 0
            ORDER BY r.created_at DESC
        '''
        cursor.execute(sql, (user_id,))
        return cursor.fetchall()
    except Error as e:
        print(f"獲取使用者需求單時發生錯誤: {e}")
        return []

def get_admin_dispatched_requirements(conn, admin_id):
    """獲取管理員已發派的需求單 (未刪除)"""
    try:
        cursor = conn.cursor()
        sql = f'''
            SELECT {_get_requirement_select_fields()}
            FROM requirements r
            {_get_requirement_joins()}
            WHERE r.assigner_id = ? AND r.is_dispatched = 1 AND r.is_deleted = 0
            ORDER BY r.created_at DESC
        '''
        cursor.execute(sql, (admin_id,))
        return cursor.fetchall()
    except Error as e:
        print(f"獲取管理員已發派需求單時發生錯誤: {e}")
        return []

def get_admin_requirements_by_staff(conn, admin_id, staff_id):
    """獲取管理員發派給特定員工的需求單 (已發派且未刪除)"""
    try:
        cursor = conn.cursor()
        sql = f'''
            SELECT {_get_requirement_select_fields()}
            FROM requirements r
            {_get_requirement_joins()}
            WHERE r.assigner_id = ? AND r.assignee_id = ? AND r.is_dispatched = 1 AND r.is_deleted = 0
            ORDER BY r.created_at DESC
        '''
        cursor.execute(sql, (admin_id, staff_id))
        return cursor.fetchall()
    except Error as e:
        print(f"依員工篩選管理員需求單時發生錯誤: {e}")
        return []

def get_admin_scheduled_requirements(conn, admin_id):
    """獲取管理員預約發派的需求單 (未發派且未刪除)"""
    try:
        cursor = conn.cursor()
        sql = f'''
            SELECT {_get_requirement_select_fields()}
            FROM requirements r
            {_get_requirement_joins()}
            WHERE r.assigner_id = ? AND r.is_dispatched = 0 AND r.is_deleted = 0
            ORDER BY r.scheduled_time ASC
        '''
        cursor.execute(sql, (admin_id,))
        return cursor.fetchall()
    except Error as e:
        print(f"獲取管理員預約需求單時發生錯誤: {e}")
        return []

def get_admin_scheduled_by_staff(conn, admin_id, staff_id):
    """獲取管理員預約給特定員工的需求單 (未發派且未刪除)"""
    try:
        cursor = conn.cursor()
        sql = f'''
            SELECT {_get_requirement_select_fields()}
            FROM requirements r
            {_get_requirement_joins()}
            WHERE r.assigner_id = ? AND r.assignee_id = ? AND r.is_dispatched = 0 AND r.is_deleted = 0
            ORDER BY r.scheduled_time ASC
        '''
        cursor.execute(sql, (admin_id, staff_id))
        return cursor.fetchall()
    except Error as e:
        print(f"依員工篩選管理員預約需求單時發生錯誤: {e}")
        return []

def dispatch_scheduled_requirements(conn):
    """檢查並發派到期的預約需求單"""
    try:
        current_time_dt = datetime.datetime.now()
        current_time_str = current_time_dt.strftime("%Y-%m-%d %H:%M:%S")
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT id FROM requirements 
            WHERE is_dispatched = 0 AND scheduled_time <= ? AND is_deleted = 0
        ''', (current_time_str,))
        
        req_ids = [row[0] for row in cursor.fetchall()]
        
        if req_ids:
            for req_id in req_ids:
                cursor.execute('''
                    UPDATE requirements 
                    SET is_dispatched = 1, created_at = ?, status = 'pending'
                    WHERE id = ?
                ''', (current_time_str, req_id)) # Use current_time_str for created_at when dispatched
            conn.commit()
            return len(req_ids)
        return 0
    except Error as e:
        print(f"自動發派預約需求單時發生錯誤: {e}")
        return 0

def has_upcoming_scheduled_requirements(conn, minutes_ahead=2):
    """檢查是否有即將到期的預約需求單（預設檢查未來2分鐘內）"""
    try:
        current_time_dt = datetime.datetime.now()
        future_time_dt = current_time_dt + datetime.timedelta(minutes=minutes_ahead)
        future_time_str = future_time_dt.strftime("%Y-%m-%d %H:%M:%S")
        current_time_str = current_time_dt.strftime("%Y-%m-%d %H:%M:%S")
        
        cursor = conn.cursor()
        cursor.execute('''
            SELECT COUNT(*) FROM requirements 
            WHERE is_dispatched = 0 AND scheduled_time > ? AND scheduled_time <= ? AND is_deleted = 0
        ''', (current_time_str, future_time_str))
        
        count = cursor.fetchone()[0]
        return count > 0
    except Error as e:
        print(f"檢查即將到期預約需求單時發生錯誤: {e}")
        return False

def cancel_scheduled_requirement(conn, req_id):
    """取消預約發派的需求單 (軟刪除)"""
    try:
        cursor = conn.cursor()
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        # Instead of deleting, mark as deleted and set status to 'cancelled' or similar
        # For now, just soft delete as per existing delete_requirement logic
        cursor.execute('''
            UPDATE requirements
            SET is_deleted = 1, deleted_at = ?, status = 'cancelled' 
            WHERE id = ? AND is_dispatched = 0 AND is_deleted = 0
        ''', (current_time, req_id))
        conn.commit()
        return cursor.rowcount > 0
    except Error as e:
        print(f"取消預約需求單時發生錯誤: {e}")
        return False

def submit_requirement(conn, req_id, comment, attachment_path=None):
    """員工提交需求單完成情況"""
    try:
        cursor = conn.cursor()
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        cursor.execute('''
            UPDATE requirements
            SET status = 'reviewing', comment = ?, completed_at = ?
            WHERE id = ? AND status = 'pending' AND is_deleted = 0
        ''', (comment, current_time, req_id))
        
        conn.commit()
        return cursor.rowcount > 0
    except Error as e:
        print(f"提交需求單時發生錯誤: {e}")
        return False

def approve_requirement(conn, req_id):
    """管理員審核通過需求單"""
    try:
        cursor = conn.cursor()
        cursor.execute('''
            UPDATE requirements
            SET status = 'completed'
            WHERE id = ? AND status = 'reviewing' AND is_deleted = 0
        ''', (req_id,))
        conn.commit()
        return cursor.rowcount > 0
    except Error as e:
        print(f"審核通過需求單時發生錯誤: {e}")
        return False

def reject_requirement(conn, req_id):
    """管理員拒絕需求單，將狀態改回未完成"""
    try:
        cursor = conn.cursor()
        # Clear previous comment and completion time when rejecting
        cursor.execute('''
            UPDATE requirements
            SET status = 'pending', comment = NULL, completed_at = NULL
            WHERE id = ? AND status = 'reviewing' AND is_deleted = 0
        ''', (req_id,))
        conn.commit()
        return cursor.rowcount > 0
    except Error as e:
        print(f"退回需求單時發生錯誤: {e}")
        return False

def invalidate_requirement(conn, req_id):
    """使需求單失效"""
    try:
        cursor = conn.cursor()
        cursor.execute('''
            UPDATE requirements
            SET status = 'invalid'
            WHERE id = ? AND is_deleted = 0
        ''', (req_id,))
        conn.commit()
        return cursor.rowcount > 0
    except Error as e:
        print(f"使需求單失效時發生錯誤: {e}")
        return False

def delete_requirement(conn, req_id):
    """刪除需求單（軟刪除）"""
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
        print(f"刪除需求單時發生錯誤: {e}")
        return False

def restore_requirement(conn, req_id):
    """恢復已刪除的需求單"""
    try:
        cursor = conn.cursor()
        # Determine a sensible status upon restoration
        # If it was 'cancelled', maybe it goes back to 'not_dispatched' if scheduled_time is in future, or 'pending'
        # For now, a simple approach: set to 'pending' if it was soft-deleted.
        # Consider more nuanced logic if needed, e.g., checking scheduled_time.
        cursor.execute('''
            UPDATE requirements
            SET is_deleted = 0, deleted_at = NULL, status = 
                CASE 
                    WHEN scheduled_time IS NOT NULL AND is_dispatched = 0 THEN 'not_dispatched'
                    ELSE 'pending'
                END
            WHERE id = ? AND is_deleted = 1
        ''', (req_id,))
        conn.commit()
        return cursor.rowcount > 0
    except Error as e:
        print(f"恢復需求單時發生錯誤: {e}")
        return False

def get_deleted_requirements(conn, admin_id):
    """獲取管理員已刪除的需求單"""
    try:
        cursor = conn.cursor()
        sql = f'''
            SELECT {_get_requirement_select_fields()}
            FROM requirements r
            {_get_requirement_joins()}
            WHERE r.assigner_id = ? AND r.is_deleted = 1
            ORDER BY r.deleted_at DESC
        '''
        cursor.execute(sql, (admin_id,))
        return cursor.fetchall()
    except Error as e:
        print(f"獲取已刪除需求單時發生錯誤: {e}")
        return [] 

# Example of how to clear all requirements (for testing, use with caution)
def clear_all_requirements_DANGEROUS(conn):
    """(危險操作) 清空所有需求單"""
    try:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM requirements")
        conn.commit()
        print(f"已清空資料庫中的所有需求單")
        return True
    except Error as e:
        print(f"清空需求單時發生錯誤: {e}")
        return False 