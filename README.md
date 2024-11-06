# Police-officer-dispatch-system
This is a police dispatch and scheduling system for police headquarter

# 警察局排班系統

這是一個用於管理警察局排班的系統，提供班表管理、人員調動、假檔分配等功能。

## 功能特點

- 排班管理
- 輪休檔次查詢
- 隊伍排序管理
- 隊內人員順序調整
- 人員資料管理
- 匯出班表報表

## 系統需求

- Python 3.8+
- MySQL 8.0+
- 必要的 Python 套件（見 requirements.txt）

## 安裝說明

1. clone專案
'''bash
git clone git@github.com:yuhao831068/Police-officer-dispatch-system.git
cd police-schedule-system
'''

2. 安裝依賴套件
'''bash
pip install -r requirements.txt
'''

3. 設定資料庫
- 在 MySQL 中執行 `database_schema.sql` 建立資料庫和表格
'''bash
mysql -u your_username -p < database_schema.sql
'''

4. 設定環境變數
- 複製 `.env.example` 為 `.env`
- 編輯 `.env` 填入你的資料庫連線資訊
'''
mysql_password=your_password
'''

## 資料庫結構

### Employee_Shift 表 (員工資料)
| 欄位          | 型別         | 說明     |
|--------------|-------------|----------|
| S_ID         | VARCHAR(10) | 警員編號(PK) |
| name         | VARCHAR(50) | 姓名     |
| team         | VARCHAR(30) | 所屬隊伍  |
| job_rank     | VARCHAR(20) | 職級     |
| current_shift| VARCHAR(10) | 目前假檔  |

### Teams 表 (隊伍資料)
| 欄位            | 型別         | 說明     |
|----------------|-------------|----------|
| team_id        | INT         | 隊伍編號(PK) |
| team_leader_id | VARCHAR(10) | 隊長編號(FK) |
| team_leader2_id| VARCHAR(10) | 第二隊長編號(FK,僅給第10隊使用) |
| member1_id     | VARCHAR(10) | 隊員1編號(FK) |
| member2_id     | VARCHAR(10) | 隊員2編號(FK) |
| member3_id     | VARCHAR(10) | 隊員3編號(FK) |
| member4_id     | VARCHAR(10) | 隊員4編號(FK) |
| member5_id     | VARCHAR(10) | 隊員5編號(FK) |
| member6_id     | VARCHAR(10) | 隊員6編號(FK) |
| member7_id     | VARCHAR(10) | 隊員7編號(FK) |

### Shift_Type 表 (班別類型)
| 欄位          | 型別         | 說明     |
|--------------|-------------|----------|
| type_name    | VARCHAR(20) | 班別名稱(PK) |
| allowed_rank | VARCHAR(20) | 允許職級  |

### Shift 表 (排班資料)
| 欄位       | 型別         | 說明     |
|-----------|-------------|----------|
| shift_id  | INT         | 班表編號(PK) |
| shift_name| VARCHAR(20) | 班別名稱  |
| S_ID      | VARCHAR(10) | 警員編號(FK) |
| shift_date| DATE        | 日期     |
| team_order| INT         | 檔排序   |
| day_order | INT         | 日排序   |

## 使用說明

1. 執行系統
'''bash
python main.py
'''

2. 主選單功能：
   - 安排今日班別
   - 安排指定日期班別
   - 查看班表
   - 查看當日輪休檔次
   - 查看隊伍排序
   - 產生空表
   - 修改班別
   - 管理隊伍人員

## 資料夾結構

'''
police-schedule-system/
├── main.py             # 主程式
├── shift_manager.py    # 班表管理類
├── database.py         # 資料庫連接管理
├── utils.py           # 工具函數
├── database_schema.sql # 資料庫結構
├── requirements.txt    # 依賴套件
├── .env.example       # 環境變數範例
└── README.md          # 說明文件
'''
