```sql
-- 建立資料庫
CREATE DATABASE IF NOT EXISTS police_schedule_db;
USE police_schedule_db;

-- 建立員工資料表
CREATE TABLE Employee_Shift (
    S_ID VARCHAR(10) PRIMARY KEY,
    name VARCHAR(50) NOT NULL,
    team VARCHAR(10) NOT NULL,
    job_rank VARCHAR(20) NOT NULL,
    current_shift VARCHAR(20) NOT NULL
);

-- 建立班表資料表
CREATE TABLE Shift (
    id INT AUTO_INCREMENT PRIMARY KEY,
    shift_name VARCHAR(20) NOT NULL,
    S_ID VARCHAR(10) NOT NULL,
    shift_date DATE NOT NULL,
    team_order INT NOT NULL,
    day_order INT NOT NULL,
    FOREIGN KEY (S_ID) REFERENCES Employee_Shift(S_ID)
);

-- 插入測試資料
INSERT INTO Employee_Shift (S_ID, name, team, job_rank, current_shift) VALUES
('C001', '李隊長', '1', '隊長', '123檔期'),
('P101', '警員101', '1', '警務員', '123檔期'),
('P102', '警員102', '1', '警務員', '456檔期'),
-- ... 其他測試資料 ...
;
```