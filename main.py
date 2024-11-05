from datetime import datetime
from shift_manager import ShiftManager
from utils import get_shifts_config, format_date


def main_menu():
    """顯示主選單"""
    print("\n=== 警察局排班系統 ===")
    print("1. 安排今日班別")
    print("2. 安排指定日期班別")
    print("3. 查看班表")
    print("4. 查看當日輪休檔次")
    print("5. 查看隊伍排序")
    print("6. 產生空表")
    print("7. 修改班別")
    print("8. 管理隊伍人員")  # 新增選項
    print("9. 退出")
    return input("請選擇功能 (1-9): ")

def handle_assign_shift(manager, shift_date, shift_name, rank):
    """處理單個班別的指派"""
    # 檢查班別是否已被分配
    is_assigned, current_emp = manager.check_shift_assigned(shift_name, shift_date)
    if is_assigned:
        print(f"\n警告：{shift_name} 目前已由 {current_emp['name']}({current_emp['team']}隊) 擔任")
        modify = input("是否要修改此班別? (y/n): ")
        if modify.lower() == 'y':
            new_sid = input(f"請輸入新的警員編號 (限{rank}): ")
            if new_sid:
                success, message = manager.modify_shift(
                    shift_name,
                    current_emp['S_ID'],
                    new_sid,
                    shift_date
                )
                print(message)
        return

    # 指派新的班別
    while True:
        s_id = input(f"\n請輸入警員編號以指派{shift_name} (限{rank}, 按Enter跳過): ")

        if not s_id:  # 跳過此班別
            break

        success, message = manager.assign_shift(shift_name, s_id, shift_date)
        print(message)

        if success:
            break
        else:
            retry = input("是否重新輸入? (y/n): ")
            if retry.lower() != 'y':
                break


def assign_shifts(manager, shift_date):
    """安排班別的主要函數"""
    print(f"\n=== 正在安排 {shift_date} 的班別 ===")

    # 先顯示當前班表
    print("\n當前班表:")
    current_shifts = manager.view_daily_shifts(shift_date)
    if current_shifts is not None:
        print(current_shifts)

    print("\n開始安排班別...")
    shifts = get_shifts_config()

    for shift_name, rank in shifts:
        handle_assign_shift(manager, shift_date, shift_name, rank)

    # 完成後顯示更新後的班表
    print("\n=== 更新後的班表 ===")
    updated_shifts = manager.view_daily_shifts(shift_date)
    if updated_shifts is not None:
        print(updated_shifts)


def handle_view_shifts(manager):
    """處理查看班表功能"""
    date_str = input("請輸入要查看的日期 (YYYY-MM-DD): ")
    try:
        shift_date = format_date(date_str)
        df = manager.view_daily_shifts(shift_date)
        if df is not None:
            print("\n=== 當日班表 ===")
            print(df)
        else:
            print("查無資料")
    except ValueError:
        print("日期格式錯誤，請使用YYYY-MM-DD格式")


def handle_check_rest_shift(manager):
    """處理查看輪休檔次功能"""
    date_str = input("請輸入要查看的日期 (YYYY-MM-DD): ")
    try:
        check_date = format_date(date_str)
        result = manager.check_current_rest_shift(check_date)
        print(result)
    except ValueError:
        print("日期格式錯誤，請使用YYYY-MM-DD格式")


def handle_view_team_orders(manager):
    """處理查看隊伍排序功能"""
    team = input("請輸入隊伍編號: ")
    date_str = input("請輸入要查看的日期 (YYYY-MM-DD): ")
    try:
        check_date = format_date(date_str)
        success, result = manager.view_team_orders(team, check_date)
        if success:
            print("\n=== 隊伍排序資訊 ===")
            print(f"隊別: {result['隊別']}")
            print(f"檔排序: {result['檔排序']}")
            print("\n人員狀態:")
            for person in result['人員狀態']:
                print(f"姓名: {person['姓名']}")
                print(f"假檔: {person['假檔']}")
                print(f"日排序: {person['日排序']}")
                print(f"狀態: {person['狀態']}")
                print("---")
        else:
            print(f"錯誤：{result}")
    except ValueError:
        print("日期格式錯誤，請使用YYYY-MM-DD格式")


def handle_generate_standby_groups(manager):
    """處理產生備勤分組功能"""
    date_str = input("請輸入要產生空表的日期 (YYYY-MM-DD): ")
    try:
        check_date = format_date(date_str)
        success, groups = manager.generate_all_standby_groups(check_date)
        if success:
            filename = manager.export_to_word(groups, check_date)
            if filename:
                print(f"已成功生成檔案：{filename}")
            else:
                print("檔案生成失敗")
        else:
            print(f"錯誤：{groups}")
    except ValueError:
        print("日期格式錯誤，請使用YYYY-MM-DD格式")


def handle_modify_shift(manager):
    """處理修改班別功能"""
    date_str = input("請輸入要修改的日期 (YYYY-MM-DD): ")
    try:
        shift_date = format_date(date_str)

        # 顯示當前班表
        print("\n當前班表:")
        df = manager.view_daily_shifts(shift_date)
        if df is not None:
            print(df)
        else:
            print("查無資料")
            return

        # 獲取要修改的資訊
        shift_name = input("\n請輸入要修改的班別: ")
        is_assigned, current_emp = manager.check_shift_assigned(shift_name, shift_date)

        if not is_assigned:
            print(f"錯誤：找不到 {shift_name} 的排班資料")
            return

        print(f"\n當前擔任 {shift_name} 的是：{current_emp['name']}({current_emp['team']}隊)")
        new_sid = input("請輸入新的警員編號: ")

        success, message = manager.modify_shift(
            shift_name,
            current_emp['S_ID'],
            new_sid,
            shift_date
        )
        print(message)

        # 顯示更新後的班表
        print("\n=== 更新後的班表 ===")
        updated_df = manager.view_daily_shifts(shift_date)
        if updated_df is not None:
            print(updated_df)

    except ValueError:
        print("日期格式錯誤，請使用YYYY-MM-DD格式")


def handle_team_management(manager):
    """處理隊伍人員管理功能"""
    while True:
        print("\n=== 隊伍人員管理 ===")
        print("1. 查看隊伍人員順序")
        print("2. 調整隊內順序")
        print("3. 更新隊員資料")
        print("4. 批次調整假檔")
        print("5. 返回主選單")

        choice = input("請選擇功能 (1-5): ")

        if choice == '1':
            team_id = input("請輸入隊伍編號: ")
            success, result = manager.view_team_member_order(team_id)
            if success:
                print(f"\n=== 第{team_id}隊人員順序 ===")
                print(result)
            else:
                print(f"錯誤：{result}")

        elif choice == '2':
            team_id = input("請輸入隊伍編號: ")
            # 先顯示當前順序
            success, result = manager.view_team_member_order(team_id)
            if not success:
                print(f"錯誤：{result}")
                continue

            print("\n當前隊伍人員順序:")
            print(result)

            print("\n請選擇調整方式：")
            print("1. 交換兩人順序")
            print("2. 重新排序")
            adjust_choice = input("請選擇 (1-2): ")

            if adjust_choice == '1':
                # 交換兩人順序
                print("\n請輸入要交換的兩個警員編號：")
                s_id1 = input("第一個警員編號: ")
                s_id2 = input("第二個警員編號: ")

                success, message = manager.swap_member_orders(team_id, s_id1, s_id2)
                print(message)

                if success:
                    # 顯示更新後的順序
                    success, updated_result = manager.view_team_member_order(team_id)
                    if success:
                        print("\n更新後的順序:")
                        print(updated_result)

            elif adjust_choice == '2':
                # 重新排序
                print("\n請依序為每位成員指定新的順序（1 開始）")
                order_changes = {}
                used_orders = set()

                for _, row in result.iterrows():
                    s_id = str(row['S_ID'])
                    name = row['姓名']
                    current_order = row['順序']

                    while True:
                        new_order = input(f"{name} 當前順序 {current_order}，新順序（按Enter保持不變）: ").strip()

                        if not new_order:  # 跳過不變的
                            break

                        try:
                            new_order = int(new_order)
                            if new_order < 1:
                                print("順序必須大於 0")
                                continue
                            if new_order in used_orders:
                                print("此順序已被使用，請選擇其他順序")
                                continue

                            order_changes[s_id] = new_order
                            used_orders.add(new_order)
                            break

                        except ValueError:
                            print("請輸入有效的數字")

                if order_changes:
                    success, message = manager.update_member_order(team_id, order_changes)
                    print(message)

                    if success:
                        # 顯示更新後的順序
                        success, updated_result = manager.view_team_member_order(team_id)
                        if success:
                            print("\n更新後的順序:")
                            print(updated_result)
                else:
                    print("未進行任何更改")

            else:
                print("無效的選擇")

        elif choice == '3':
            s_id = input("請輸入警員編號: ")
            print("\n要更新什麼資料？")
            print("1. 調整隊別")
            print("2. 調整假檔")
            print("3. 同時調整")
            update_choice = input("請選擇 (1-3): ")

            new_team = None
            new_shift = None

            if update_choice in ['1', '3']:
                new_team = input("請輸入新的隊別: ")
            if update_choice in ['2', '3']:
                print("\n可用的假檔：")
                print("1. 123檔期")
                print("2. 456檔期")
                print("3. 789檔期")
                shift_choice = input("請選擇新的假檔 (1-3): ")
                shift_map = {'1': '123檔期', '2': '456檔期', '3': '789檔期'}
                new_shift = shift_map.get(shift_choice)

            success, message = manager.update_team_member(s_id, new_team, new_shift)
            print(message)

        elif choice == '4':
            team_id = input("請輸入隊伍編號: ")
            success, result = manager.view_team_member_order(team_id)
            if not success:
                print(f"錯誤：{result}")
                continue

            print("\n當前隊伍成員:")
            print(result)

            print("\n請依序為每位成員指定新的假檔")
            print("可用的假檔：1=123檔期, 2=456檔期, 3=789檔期")

            shift_assignments = {}
            shift_map = {'1': '123檔期', '2': '456檔期', '3': '789檔期'}

            for _, row in result.iterrows():
                s_id = str(row['S_ID'])
                name = row['姓名']
                shift_choice = input(f"請為 {name} 選擇新的假檔 (1-3，按Enter跳過): ")
                if shift_choice in shift_map:
                    shift_assignments[s_id] = shift_map[shift_choice]

            if shift_assignments:
                success, message = manager.bulk_update_team_shifts(team_id, shift_assignments)
                print(message)
            else:
                print("未進行任何更改")

        elif choice == '5':
            break

        else:
            print("無效的選擇，請重新輸入")


def handle_order_adjustment(manager, team_id):
    """處理順序調整功能"""
    try:
        # 先顯示當前順序
        success, result = manager.view_team_member_order(team_id)
        if not success:
            print(f"錯誤：{result}")
            return

        print("\n當前隊伍人員順序:")
        print(result)

        print("\n請依序為每位成員指定新的順序（1 開始）")
        order_changes = {}
        used_orders = set()

        for _, row in result.iterrows():
            s_id = str(row['S_ID'])  # 確保 S_ID 是字符串
            name = row['姓名']
            current_order = row['順序']

            while True:
                new_order = input(f"{name} 當前順序 {current_order}，新順序（按Enter保持不變）: ").strip()

                if not new_order:  # 跳過不變的
                    break

                try:
                    new_order = int(new_order)
                    if new_order < 1:
                        print("順序必須大於 0")
                        continue
                    if new_order in used_orders:
                        print("此順序已被使用，請選擇其他順序")
                        continue

                    order_changes[s_id] = new_order
                    used_orders.add(new_order)
                    break

                except ValueError:
                    print("請輸入有效的數字")

        if order_changes:
            success, message = manager.update_member_order(team_id, order_changes)
            print(message)

            if success:
                # 顯示更新後的順序
                success, updated_result = manager.view_team_member_order(team_id)
                if success:
                    print("\n更新後的順序:")
                    print(updated_result)
        else:
            print("未進行任何更改")

    except Exception as err:
        print(f"處理過程中發生錯誤: {str(err)}")

def main():
    """主程式入口"""
    manager = ShiftManager()

    try:
        manager.connect()
        print("歡迎使用警察局排班系統")

        while True:
            try:
                choice = main_menu()

                if choice == '1':
                    shift_date = datetime.now().date()
                    assign_shifts(manager, shift_date)

                elif choice == '2':
                    date_str = input("請輸入日期 (YYYY-MM-DD): ")
                    try:
                        shift_date = format_date(date_str)
                        assign_shifts(manager, shift_date)
                    except ValueError:
                        print("日期格式錯誤，請使用YYYY-MM-DD格式")

                elif choice == '3':
                    handle_view_shifts(manager)

                elif choice == '4':
                    handle_check_rest_shift(manager)

                elif choice == '5':
                    handle_view_team_orders(manager)

                elif choice == '6':
                    handle_generate_standby_groups(manager)

                elif choice == '7':
                    handle_modify_shift(manager)

                elif choice == '8':

                    handle_team_management(manager)

                elif choice == '9':

                    print("感謝使用，再見！")
                    break

                else:
                    print("無效的選擇，請重新輸入")

            except Exception as e:
                print(f"操作過程中發生錯誤: {str(e)}")
                print("請重試或聯繫系統管理員")
                continue

    except Exception as e:
        print(f"系統啟動錯誤: {str(e)}")

    finally:
        manager.disconnect()

        print("系統已安全關閉")


if __name__ == "__main__":
    main()