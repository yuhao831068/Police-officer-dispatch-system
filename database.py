import os
import mysql.connector
from dotenv import load_dotenv

load_dotenv()


class DatabaseConnection:
    """管理資料庫連接的類"""

    def __init__(self):
        """初始化資料庫配置"""
        self.db_config = {
            'host': 'localhost',
            'user': 'root',
            'password': os.getenv("mysql_password"),
            'database': 'police_schedule_db'
        }
        self.conn = None
        self.cursor = None

    def connect(self):
        """建立資料庫連接"""
        try:
            self.conn = mysql.connector.connect(**self.db_config)
            self.cursor = self.conn.cursor()
            print("資料庫連接成功")
        except mysql.connector.Error as err:
            print(f"資料庫連接錯誤: {err}")
            raise

    def disconnect(self):
        """關閉資料庫連接"""
        try:
            if self.cursor:
                self.cursor.close()
            if self.conn:
                self.conn.close()
            print("資料庫連接已關閉")
        except mysql.connector.Error as err:
            print(f"關閉資料庫連接時發生錯誤: {err}")

    def get_cursor(self):
        """獲取資料庫游標"""
        return self.cursor

    def get_connection(self):
        """獲取資料庫連接"""
        return self.conn