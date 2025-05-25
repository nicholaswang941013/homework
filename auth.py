from database import create_connection, get_user_by_username
from models import User


def login(username, password):
    conn = create_connection()
    if conn is not None:
        user = get_user_by_username(conn, username)
        conn.close()

        if user is not None and user[2] == password:  # 檢查密碼
            return {
                "success": True,
                "user_info": User(
                    id=user[0],
                    username=user[1],
                    name=user[3],
                    email=user[4],
                    role=user[5]
                )
            }
    return {
        "success": False,
        "message": "使用者名稱或密碼錯誤"
    } 