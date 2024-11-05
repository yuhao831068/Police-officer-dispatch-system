from datetime import datetime


def get_team_order(team, current_month):
    """
    取得隊伍在當前月份的排序

    Args:
        team: 隊伍編號
        current_month: 當前月份

    Returns:
        int: 排序順序(1-3)
    """
    # 基礎排序
    base_orders = {
        # 123檔
        '1': 2, '2': 3, '3': 1,
        # 456檔
        '4': 2, '5': 3, '6': 1,
        # 789檔
        '7': 2, '8': 3, '9': 1,
        # 其他隊
        '11': 1, '13': 2, '14': 3
    }

    order = base_orders.get(str(team), 0)

    # 計算當月是第幾個循環月
    rotation_month = (current_month - 1) % 3 + 1

    # 根據循環月調整順序
    if rotation_month != 1:
        if order == 3:
            order = 2 if rotation_month == 2 else 1
        elif order == 2:
            order = 1 if rotation_month == 2 else 3
        elif order == 1:
            order = 3 if rotation_month == 2 else 2

    return order


def format_date(date_input):
    """
    統一日期格式轉換

    Args:
        date_input: 日期輸入(可以是字符串或datetime對象)

    Returns:
        date: 標準化的日期對象
    """
    if isinstance(date_input, str):
        return datetime.strptime(date_input, '%Y-%m-%d').date()
    elif isinstance(date_input, datetime):
        return date_input.date()
    return date_input


def get_shifts_config():
    """
    獲取班別配置

    Returns:
        list: 包含(班別名稱, 職級要求)的元組列表
    """
    return [
        # 警務員的班別
        ('A班', '警務員'),
        ('B班', '警務員'),
        ('C班', '警務員'),
        ('D班', '警務員'),
        ('E班', '警務員'),
        ('上值日', '警務員'),
        ('下值日', '警務員'),
        ('日勤務管理員', '警務員'),
        ('夜勤務管理員', '警務員'),
        ('日械彈管理員', '警務員'),
        ('夜械彈管理員', '警務員'),

        # 隊長的班別
        ('日值日官', '隊長'),
        ('夜值日官', '隊長'),

        # 副大隊長的班別
        ('值班副大隊長', '副大隊長')
    ]


def get_rank_restrictions():
    """
    獲取職級限制配置

    Returns:
        dict: 班別對應的職級要求
    """
    return {
        'A班': '警務員',
        'B班': '警務員',
        'C班': '警務員',
        'D班': '警務員',
        'E班': '警務員',
        '上值日': '警務員',
        '下值日': '警務員',
        '日值日官': '隊長',
        '夜值日官': '隊長',
        '值班副大隊長': '副大隊長',
        '日勤務管理員': '警務員',
        '夜勤務管理員': '警務員',
        '日械彈管理員': '警務員',
        '夜械彈管理員': '警務員'
    }