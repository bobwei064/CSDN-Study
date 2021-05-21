import sys
import time
import os
import re

from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QTime

from My_Windows.AddPlayerWindow import AddPlayerWindow
from My_Windows.AcctChooseWindow import AcctChooseWindow
from My_Windows.AccountEntryWindow import AcctEntryWindow
from My_Windows.rollingAcctChooseWindow import RollingAcctChooseWindow
from My_Windows.RollingEntryWindow import RollingEntryWindow
from My_Windows.DailyCashWindow import DailyCashWindow
from My_tools.Excel_Tools import *
from My_Windows.MainWindow import AppWindow
from My_tools.MessageSender import send_message
from My_tools.PopUpMessage import *


def main():
    app = QApplication(sys.argv)

    # 实例化需要的界面
    app_window = AppWindow()
    add_player_window = AddPlayerWindow()
    acct_choose_window = AcctChooseWindow()
    acct_entry_window = AcctEntryWindow()
    rolling_acct_choose_window = RollingAcctChooseWindow()
    entry_rolling_window = RollingEntryWindow()
    daily_cash_window = DailyCashWindow()

    # 显示主界面
    app_window.show()

    # 显示现金日记账窗口的函数
    def open_daily_cash():
        daily_cash_window.show()
        app_window.close()

    # 关闭现金日记账窗口的函数
    def close_daily_cash():
        daily_cash_window.amount_change_lineEdit.setText("")
        daily_cash_window.remark_plainTextEdit.clear()
        daily_cash_window.close()
        app_window.show()

    daily_cash_window.cancel_Button.clicked.connect(close_daily_cash)

    # 绑定主界面的现金日记账显示按钮
    app_window.otherButton.clicked.connect(open_daily_cash)

    # 绑定主页面的客户录入按钮的函数，显示录入界面，隐藏主界面
    def show_add_player_window():
        # 先获取可用的账户号，添加到界面中
        account_index, account_number = get_new_player_index_and_number()
        add_player_window.acctNumberEdit.setText(account_number)
        # 显示添加用户界面
        add_player_window.show()
        # 隐藏主界面
        app_window.hide()

    app_window.newPlayerButton.clicked.connect(show_add_player_window)

    # 清空客户信息录入界面的函数
    def clear_player_info():
        # 退出录入界面的时候把数据清空
        add_player_window.acctNameEdit.setText("")
        add_player_window.phoneEdit.setText("")
        add_player_window.playerCardEdit.setText("")
        add_player_window.remarkEdit.setText("")
        add_player_window.playerNameEdit.setText("")
        add_player_window.loadLimitEdit.setText("0.0")
        add_player_window.settlementRateEdit.setText("0.0")
        add_player_window.acctNumberEdit.setText("")

    # 客户录入界面退出按钮函数及绑定
    def exit_add_player_window():
        # 退出录入界面的时候把数据清空
        clear_player_info()
        # 关闭录入界面
        add_player_window.close()
        # 显示主界面
        app_window.show()

    add_player_window.quitButton.clicked.connect(exit_add_player_window)

    # 客户录入界面的保存按钮函数及绑定 已完成
    def save_new_player():

        # 把要保存在客户基本信息表的数据获取出来
        acct_index, nothing_important = get_new_player_index_and_number()
        acct_name = add_player_window.acctNameEdit.text()
        acct_number = add_player_window.acctNumberEdit.text()
        player_name = add_player_window.playerNameEdit.text()
        phone = add_player_window.phoneEdit.text()
        currency_type = add_player_window.currencyTypecomboBox.currentText()
        player_card = add_player_window.playerCardEdit.text()
        loan_limit = add_player_window.loadLimitEdit.text()
        settlement_type = add_player_window.settlementTypecomboBox.currentText()
        settlement_rate = add_player_window.settlementRateEdit.text()
        remark = add_player_window.remarkEdit.text()
        # 序号、账户名和账户号不可以为空，如果为空弹出提示中断进程
        if not acct_name or not acct_number:
            # 这里要写弹出提示框的代码
            show_warning(add_player_window, "户名和账号不允许为空！")
            return
        try:
            loan_limit = float(loan_limit)
            settlement_rate = float(settlement_rate)
        except:
            pass

        # 把数据封装成列表，方便写入表格中
        info = [acct_index, acct_name, acct_number, player_name, phone, currency_type, player_card, loan_limit,
                settlement_type, settlement_rate, remark]

        # 调用保存数据的函数把数据保存到表格中
        player_info_result = save_player_info_book(info, player_info_file_name)
        # 保存成功给提示框
        show_info(add_player_window, "%s表格已更新并保存成功！" % player_info_result)
        show_add_player_window()

        # 把要保存到对应总账的初始余额，存取款数和当前余额先列出来
        opening_balance = 0.0
        opening_amount = 0.0
        current_balance = opening_balance + opening_amount

        # 把除了序号之外的数据，放入列表中，要保存之前，再把序号插入到列表首位即可
        data_to_be_saved_in_acct_file = [acct_name, acct_number, opening_balance, opening_amount, current_balance]

        # 关联币种的要写入对应的总账中去
        if currency_type == 'PHP':
            # 先获取PHP总账的最新序号
            php_workbook, php_sheet = get_workbook(php_acct_file_name)
            # 表头占了3行，要减掉
            entry_index = php_sheet.max_row - 3

            data_to_be_saved_in_acct_file.insert(0, entry_index)
            php_result = save_work_book(data_to_be_saved_in_acct_file, php_acct_file_name)
            # 保存成功给提示框
            show_info(add_player_window, "%s表格已更新并保存成功！" % php_result)

            # PHP账户则需要再关联洗码和上下水总账中去
            # 关联保存入洗码总账中
            rolling_workbook, rolling_sheet = get_workbook(total_rolling_file_name)
            rolling_index = rolling_sheet.max_row - 4

            rolling_remark = ""
            rolling_data = [rolling_index, acct_name, acct_number, opening_balance, opening_amount, current_balance,
                            rolling_remark]
            rolling_result = save_work_book(rolling_data, total_rolling_file_name)
            # 保存成功给提示框
            show_info(add_player_window, "%s表格已更新并保存成功！" % rolling_result)

            # 关联上下水总账
            win_lose_workbook, win_lose_sheet = get_workbook(win_lose_total_file_name)
            win_lose_index = win_lose_sheet.max_row - 4

            win_lose_remark = ""
            win_lose_data = [win_lose_index, acct_name, acct_number, opening_balance, opening_amount, current_balance,
                             win_lose_remark]
            result = save_work_book(win_lose_data, win_lose_total_file_name)
            # 保存成功给提示框
            show_info(add_player_window, "%s表格已更新并保存成功！" % result)

        if currency_type == "HKD":
            # 先获取PHP总账的最新序号
            hkd_workbook, hkd_sheet = get_workbook(hkd_acct_file_name)
            # 表头占了3行，要减掉
            entry_index = hkd_sheet.max_row - 5

            data_to_be_saved_in_acct_file.insert(0, entry_index)
            result = save_work_book(data_to_be_saved_in_acct_file, hkd_acct_file_name)
            # 保存成功给提示框
            show_info(add_player_window, "%s表格已更新并保存成功！" % result)

        if currency_type == "USD":
            # 先获取PHP总账的最新序号
            usd_workbook, usd_sheet = get_workbook(usd_acct_file_name)
            # 表头占了3行，要减掉
            entry_index = usd_sheet.max_row - 4

            data_to_be_saved_in_acct_file.insert(0, entry_index)
            result = save_work_book(data_to_be_saved_in_acct_file, usd_acct_file_name)
            # 保存成功给提示框
            show_info(add_player_window, "%s表格已更新并保存成功！" % result)

        if currency_type == "RMB":
            # 先获取PHP总账的最新序号
            rmb_workbook, rmb_sheet = get_workbook(rmb_acct_file_name)
            # 表头占了3行，要减掉
            entry_index = rmb_sheet.max_row - 3

            data_to_be_saved_in_acct_file.insert(0, entry_index)
            result = save_work_book(data_to_be_saved_in_acct_file, rmb_acct_file_name)
            # 保存成功给提示框
            show_info(add_player_window, "%s表格已更新并保存成功！" % result)

        else:
            pass

        # 在账目明细里也要插入一个对应的数据
        # 先获取日期和时间的文本
        temp = time.strftime("%Y-%m-%d %H:%M", time.localtime()).split(' ')
        temp_date, entry_time = temp
        d = temp_date.split('-')
        entry_date = d[0] + "年" + d[1] + "月" + d[2] + "日"
        acct_summary_workbook, acct_summary_sheet = get_workbook(acct_summary_file_name)
        entry_index = acct_summary_sheet.max_row - 1
        handling_staff = "系统录入"
        summary_remark = "客户新开账户，0余额入账"
        summary_data = [entry_index, entry_date, entry_time, acct_name, acct_number, player_name, opening_balance,
                        opening_amount, current_balance, handling_staff, summary_remark]

        result = save_work_book(summary_data, acct_summary_file_name)
        # 保存成功给提示框
        show_info(add_player_window, "%s表格已更新并保存成功！" % result)

        # 然后要发送一个账目更新信息
        # 构造要发送的信息
        opening_message = """
德胜厅客户账目更新    
    
日期：  %s  
时间：  %s  
    
户名：  %s  
账号：  %s  
    
之前余额：  %f   万
    
本次存/取款：  %f   万
存/取款人：  %s 
存/取款备注：  %s  
    
当前账户余额：  %f   万
    
跟单员工：  %s  
备注：    请仔细核对，如有错漏，以账房数据为准！
""" % (entry_date, entry_time, acct_name, acct_number, opening_balance, opening_amount, player_name, summary_remark,
       current_balance, handling_staff)
        # 发送开户信息
        send_message(opening_message)
        show_info(add_player_window, "新开户账目信息已发送财务群！")

        # 保存成功后给出提示，并清空原来输入的内容，隐藏录入界面，显示主界面
        clear_player_info()
        add_player_window.close()
        app_window.show()

    # 绑定录入界面的保存按钮
    add_player_window.saveButton.clicked.connect(save_new_player)

    # 绑定主界面的账目明细、洗码明细、上下水明细按钮，用Excel打开对应的表格
    app_window.detailAccountButton.clicked.connect(lambda x: open_workbook(acct_summary_file_name))
    app_window.detailRollingButton.clicked.connect(lambda x: open_workbook(rolling_summary_file_name))
    app_window.detailWinLoseButton.clicked.connect(lambda x: open_workbook(win_lose_summary_file_name))

    # 绑定主界面的PHP总账、USD总账、HKD总账、RMB总账、洗码总账、上下水总账按钮，用Excel打开对应的表格
    app_window.phpSummaryButton.clicked.connect(lambda x: open_workbook(php_acct_file_name))
    app_window.usdSummaryButton.clicked.connect(lambda x: open_workbook(usd_acct_file_name))
    app_window.hkdSummaryButton.clicked.connect(lambda x: open_workbook(hkd_acct_file_name))
    app_window.rmbSummaryButton.clicked.connect(lambda x: open_workbook(rmb_acct_file_name))
    app_window.rollingSummaryButton.clicked.connect(lambda x: open_workbook(total_rolling_file_name))
    app_window.winLoseSummaryButton.clicked.connect(lambda x: open_workbook(win_lose_total_file_name))

    # 绑定主界面的客户查询按钮，用Excel打开对应的表格
    app_window.playerCheckButton.clicked.connect(lambda x: open_workbook(player_info_file_name))

    # 绑定主界面的上水客户和下水客户按钮
    app_window.winPlayerButton.clicked.connect(lambda x: get_win_lose_players("win"))
    app_window.losePlayerButton.clicked.connect(lambda x: get_win_lose_players("lose"))

    # 绑定主界面的欠款查询按钮，用Excel打开对应的表格
    app_window.loanCheckButton.clicked.connect(get_loan_list)

    # 绑定主界面账目录入按钮函数
    def show_acct_choose_window():
        acct_choose_window.show()
        app_window.close()

    app_window.newAccountingButton.clicked.connect(show_acct_choose_window)

    # 清空账户选择界面
    def clear_acct_chosen():
        acct_choose_window.acct_name_combobox.setCurrentIndex(0)
        acct_choose_window.lineEdit.setText("")

    # 清空账目录入界面信息，为下一条录入准备
    def clear_acct_entry_for_next_entry():
        acct_entry_window.date_lineEdit.setText("")
        acct_entry_window.acct_name_lineEdit.setText("")
        acct_entry_window.acct_number_lineEdit.setText("")
        acct_entry_window.balance_before_lineEdit.setText("")
        acct_entry_window.amount_change_lineEdit.setText("")
        acct_entry_window.Current_balance_lineEdit.setText("")

    # 完全清空账目录入界面信息
    def clear_acct_entry():
        acct_entry_window.date_lineEdit.setText("")
        acct_entry_window.acct_name_lineEdit.setText("")
        acct_entry_window.acct_number_lineEdit.setText("")
        acct_entry_window.balance_before_lineEdit.setText("")
        acct_entry_window.amount_change_lineEdit.setText("")
        acct_entry_window.Current_balance_lineEdit.setText("")
        acct_entry_window.depositor_lineEdit.setText("")
        acct_entry_window.handling_staff_lineEdit.setText("")
        acct_entry_window.remark_lineEdit.setText("")

    # 关闭账目录入界面的函数
    def close_acct_entry():
        clear_acct_entry()
        acct_entry_window.close()
        acct_choose_window.close()
        app_window.show()

    # 继续做下一条账目的函数
    def next_acct_entry():
        clear_acct_entry_for_next_entry()
        clear_acct_chosen()
        acct_entry_window.close()
        acct_choose_window.show()

    # 绑定账目录入界面关闭的函数
    acct_entry_window.pushButton_2.clicked.connect(close_acct_entry)

    # 绑定账户选择界面的退出按钮
    def close_acct_choose_window():
        # 退出界面的时候，清空已经选择的界面
        clear_acct_chosen()
        acct_choose_window.close()
        app_window.show()

    acct_choose_window.cancelButton.clicked.connect(close_acct_choose_window)

    # 用于设置账目录入界面的日期和时间的函数
    def set_entry_date_time():
        entry_date_time = time.strftime("%Y年%m月%d日 %H:%M", time.localtime()).split(' ')
        acct_entry_window.timeEdit.setTime(QTime.currentTime())
        entry_date = entry_date_time[0]
        entry_time = entry_date_time[1]
        # 返回日期和时间的str格式，用于存入表格和发送信息
        return entry_date, entry_time

    # 用于设置洗码录入界面的日期和时间的函数
    def set_entry_date_time_for_rolling():
        entry_date_time = time.strftime("%Y年%m月%d日 %H:%M", time.localtime()).split(' ')
        entry_rolling_window.start_time_timeEdit.setTime(QTime.currentTime())
        entry_rolling_window.end_time_timeEdit.setTime(QTime.currentTime())
        entry_date = entry_date_time[0]
        entry_time = entry_date_time[1]
        # 返回日期和时间的str格式，用于存入表格和发送信息
        return entry_date, entry_time

    # 账户选择界面的确认键函数 已完成
    def confirm_acct():
        # 获取账户名字典，用于提取账号，来查找余额
        acct_dict = get_acct_by_currency_and_acct_num()

        # 确认账号是否选择好
        flag = acct_choose_window.acct_name_combobox.currentIndex()
        amount_change = acct_choose_window.lineEdit.text()

        # 如果账户没有选择，弹出提示并中断进程
        if not flag or not amount_change:
            show_warning(acct_choose_window, "请输入或选择要录入的账户和存取款金额！")
            return
        else:
            # 设置入账日期和时间
            entry_date, entry_time = set_entry_date_time()

            acct_name = acct_choose_window.acct_name_combobox.currentText()

            # 从字典中获取账号
            acct_number = acct_dict[acct_name][0]
            # 获取该账户的余额
            acct_balance = get_acct_balance(acct_name)

            # 把输入的存取款数转化为数字
            amount_change = round(float(amount_change))
            amount_change_str = format(amount_change, ',')
            acct_entry_window.amount_change_lineEdit.setText(amount_change_str)

            # 计算出当前余额，并存入录入界面中
            balance_now = acct_balance + amount_change
            acct_entry_window.Current_balance_lineEdit.setText(format(balance_now, ','))

        # 把账户的信息填入到账目录入界面中
        acct_entry_window.date_lineEdit.setText(entry_date)
        acct_entry_window.acct_name_lineEdit.setText(acct_name)
        acct_entry_window.acct_number_lineEdit.setText(acct_number)
        acct_entry_window.balance_before_lineEdit.setText(format(acct_balance, ','))

        acct_entry_window.show()
        clear_acct_chosen()
        acct_choose_window.close()

    # 账目录入界面的保存函数 已完成
    def save_acct_change():
        # 获取客户字典
        acct_dict = get_acct_by_currency_and_acct_num()
        # 先获取需要的数据
        entry_date = acct_entry_window.date_lineEdit.text()
        entry_time = acct_entry_window.timeEdit.text()
        acct_name = acct_entry_window.acct_name_lineEdit.text()

        # 根据户名获取账户类型，以便对相应币种的总账做出修改
        acct_type = acct_dict[acct_name][1]

        acct_number = acct_entry_window.acct_number_lineEdit.text()
        depositor = acct_entry_window.depositor_lineEdit.text()
        if depositor == "":
            show_warning(acct_entry_window, "存/取款人未输入！")
            return

        balance_before = round(float(acct_entry_window.balance_before_lineEdit.text().replace(',', '')), 8)
        amount_change = acct_entry_window.amount_change_lineEdit.text().replace(',', '')

        # 转成数字格式
        amount_change = round(float(amount_change), 8)

        new_balance_str = acct_entry_window.Current_balance_lineEdit.text()
        new_balance = round(float(new_balance_str.replace(',', '')), 8)

        staff = acct_entry_window.handling_staff_lineEdit.text()
        if staff == "":
            show_warning(acct_entry_window, "跟单员工未输入！")
            return
        remark = acct_entry_window.remark_lineEdit.text()
        if remark == "":
            show_warning(acct_entry_window, "备注未输入！")
            return

        # 先构造数据，保存到账目明细库中
        # 账目明细库的最后有效行row来计算序号，账目明细库要-1
        detail_index = get_sheet_row(acct_summary_file_name) - 1
        detail_data = [detail_index, entry_date, entry_time, acct_name, acct_number, depositor, balance_before,
                       amount_change, new_balance, staff, remark]

        # 保存账目明细更新，成功则给出提示
        detail_result = save_work_book(detail_data, acct_summary_file_name)
        if detail_result:
            show_info(acct_entry_window, "%s已更新保存！" % detail_result)
        else:
            show_warning(acct_entry_window, "%s无法保存更新，请手动在对应工作表填入数据！" % detail_result)
            return

        # 根据币种，把数据保存到相应的总账中
        if acct_type == 'PHP':
            update_file = php_acct_file_name
        elif acct_type == 'USD':
            update_file = usd_acct_file_name
        elif acct_type == 'RMB':
            update_file = rmb_acct_file_name
        else:
            update_file = hkd_acct_file_name

        # 查找对应账户的存取款数和账户余额，更新该调账目
        workbook, sheet = get_workbook(update_file)

        for row in range(5, sheet.max_row + 1):
            target_cell = sheet.cell(row=row, column=2)
            target_acct = target_cell.value
            if acct_name == target_acct:
                target_row = target_cell.row
                # 更新存取款数
                sheet.cell(row=target_row, column=5, value=new_balance)
                # 更新账户余额
                sheet.cell(row=target_row, column=6, value=new_balance)
                break

        # 保存总账更新，成功则给出提示
        try:
            workbook.save(update_file)
            show_info(acct_entry_window, "%s已更新保存！" % update_file.split("/")[-1])
        except:
            show_warning(acct_entry_window, "%s无法保存更新，请手动在对应工作表填入数据！" % update_file.split("/")[-1])
            return

        # 构造信息并发送信息
        update_message = """
德胜厅客户账目更新

日期：  %s
时间：  %s

户名：  %s
账号：  %s

之前余额：  %s   万

本次存/取款：  %s   万
存/取款人：  %s
存/取款备注：  %s

当前账户余额：  %s   万

跟单员工：  %s
备注：    请仔细核对，如有错漏，以账房数据为准！
""" % (entry_date, entry_time, acct_name, acct_number, format(balance_before, ','), format(amount_change, ','),
       depositor, remark, new_balance_str, staff)

        # 发送信息
        send_message(update_message)
        show_info(acct_entry_window, "账目信息已发送财务群！")

        # 询问是否需要继续录入
        continue_flag = ask_confirmation(acct_entry_window, "是否继续录入下一条账目？")
        if continue_flag:
            next_acct_entry()
        else:
            close_acct_entry()

    # 绑定录入界面的保存按钮
    acct_entry_window.pushButton_3.clicked.connect(save_acct_change)

    # 打开洗码选择账户界面的函数
    def show_rolling_acct_choose():
        rolling_acct_choose_window.show()
        app_window.close()

    app_window.newRollingButton.clicked.connect(show_rolling_acct_choose)

    # 关闭洗码选择账户界面的函数
    def close_rolling_acct_choose():
        rolling_acct_choose_window.acct_name_combobox.setCurrentIndex(0)
        rolling_acct_choose_window.borrow_lineEdit.setText("")
        rolling_acct_choose_window.return_lineEdit.setText("")
        rolling_acct_choose_window.withdraw_lineEdit.setText("")
        rolling_acct_choose_window.deposit_lineEdit.setText("")
        rolling_acct_choose_window.actual_rolling_lineEdit.setText("")
        rolling_acct_choose_window.settled_rolling_lineEdit.setText("")

        rolling_acct_choose_window.close()
        app_window.show()

    rolling_acct_choose_window.cancelButton.clicked.connect(close_rolling_acct_choose)

    # 主界面洗码录入的函数 已完成
    def start_rolling():
        # 获取账户名字典，用于提取账号，来查找余额
        acct_dict = get_acct_by_currency_and_acct_num()

        # 确认账号是否选择好
        flag = rolling_acct_choose_window.acct_name_combobox.currentIndex()
        # 确认洗码数据和上下水数据是否填好
        this_borrow_str = rolling_acct_choose_window.borrow_lineEdit.text()
        this_return_str = rolling_acct_choose_window.return_lineEdit.text()
        this_cashout_str = rolling_acct_choose_window.withdraw_lineEdit.text()
        this_deposit_str = rolling_acct_choose_window.deposit_lineEdit.text()
        this_rolling_str = rolling_acct_choose_window.actual_rolling_lineEdit.text()
        this_settled_str = rolling_acct_choose_window.settled_rolling_lineEdit.text()

        # 如果账户没有选择，弹出提示并中断进程
        if not flag or not this_borrow_str or not this_return_str or not this_cashout_str or not this_deposit_str or not this_rolling_str or not this_settled_str:
            show_warning(acct_choose_window, "请输入或选择要录入的账户和所需数据！")
            return
        else:
            # 设置入账日期和时间
            entry_date, entry_time = set_entry_date_time_for_rolling()

            acct_name = rolling_acct_choose_window.acct_name_combobox.currentText()

            # 从字典中获取账号
            acct_number = acct_dict[acct_name][0]
            # 获取该账户的余额
            before_balance = get_acct_balance(acct_name)
            # 获取该账户的上下水总额和洗码总额
            win_lose_before, rolling_before = get_amount_of_win_lose_and_rolling(acct_name)
            if not win_lose_before or not rolling_before:
                show_warning(acct_choose_window, "该账户没有洗码或上下水数据，请把账户信息手动添加到对应的总账中再重试！")
                return
            else:
                # 把需要计算的数据转换成数字
                this_borrow = round(float(this_borrow_str), 8)
                this_return = round(float(this_return_str), 8)
                this_cashout = round(float(this_cashout_str), 8)
                this_deposit = round(float(this_deposit_str), 8)
                this_rolling = round(float(this_rolling_str), 8)
                this_settled = round(float(this_settled_str), 8)

                # 计算需要的数据
                this_win = this_return - this_borrow
                balance_now = before_balance + this_deposit
                rolling_total = rolling_before + this_rolling
                win_total = win_lose_before + this_win

                # 把数据填充到洗码录入界面中
                entry_rolling_window.entry_date_lineEdit.setText(entry_date)
                entry_rolling_window.acct_name_lineEdit.setText(acct_name)
                entry_rolling_window.acct_number_lineEdit.setText(acct_number)
                entry_rolling_window.before_balance_lineEdit.setText(format(before_balance, ','))
                entry_rolling_window.borrow_lineEdit.setText(format(this_borrow, ','))
                entry_rolling_window.return_lineEdit.setText(format(this_return, ','))
                entry_rolling_window.this_win_lineEdit.setText(format(this_win, ','))
                entry_rolling_window.cashout_lineEdit.setText(format(this_cashout, ','))
                entry_rolling_window.deposit_lineEdit.setText(format(this_deposit, ','))
                entry_rolling_window.current_balance_lineEdit.setText(format(balance_now, ','))
                entry_rolling_window.this_rolling_lineEdit.setText(format(this_rolling, ','))
                entry_rolling_window.settled_rolling_lineEdit.setText(format(this_settled, ','))
                entry_rolling_window.total_rolling_lineEdit.setText(format(rolling_total, ','))
                entry_rolling_window.total_win_lineEdit.setText(format(win_total, ','))

        # 打开洗码录入界面并关闭账户选择界面
        entry_rolling_window.show()
        rolling_acct_choose_window.close()

    # 绑定洗码账户选择界面的确认按钮
    rolling_acct_choose_window.confirmButton.clicked.connect(start_rolling)

    # 洗码录入界面退出函数 已完成
    def close_rolling_entry():
        entry_rolling_window.entry_date_lineEdit.setText("")
        entry_rolling_window.acct_name_lineEdit.setText("")
        entry_rolling_window.acct_number_lineEdit.setText("")
        entry_rolling_window.before_balance_lineEdit.setText("")
        entry_rolling_window.borrow_lineEdit.setText("")
        entry_rolling_window.return_lineEdit.setText("")
        entry_rolling_window.this_win_lineEdit.setText("")
        entry_rolling_window.cashout_lineEdit.setText("")
        entry_rolling_window.deposit_lineEdit.setText("")
        entry_rolling_window.current_balance_lineEdit.setText("")
        entry_rolling_window.this_rolling_lineEdit.setText("")
        entry_rolling_window.settled_rolling_lineEdit.setText("")
        entry_rolling_window.total_rolling_lineEdit.setText("")
        entry_rolling_window.total_win_lineEdit.setText("")
        entry_rolling_window.player_lineEdit.setText("")
        entry_rolling_window.staff_lineEdit.setText("")
        entry_rolling_window.rolling_remark_plainTextEdit.clear()
        entry_rolling_window.win_remark_plainTextEdit.clear()

        close_rolling_acct_choose()
        entry_rolling_window.close()
        app_window.show()

    # 绑定洗码录入界面的退出按钮
    entry_rolling_window.cancel_pushButton.clicked.connect(close_rolling_entry)

    # 保存洗码和上下水数据的函数 已完成
    def rolling_update():
        # 从录入界面获取数据
        entry_date = entry_rolling_window.entry_date_lineEdit.text()
        start_time = entry_rolling_window.start_time_timeEdit.text()
        end_time = entry_rolling_window.end_time_timeEdit.text()
        acct_name = entry_rolling_window.acct_name_lineEdit.text()
        acct_number = entry_rolling_window.acct_number_lineEdit.text()

        # 要转换成数字的数据
        old_balance_str = entry_rolling_window.before_balance_lineEdit.text()
        this_borrow_str = entry_rolling_window.borrow_lineEdit.text()
        this_return_str = entry_rolling_window.return_lineEdit.text()
        this_win_str = entry_rolling_window.this_win_lineEdit.text()
        this_cashout_str = entry_rolling_window.cashout_lineEdit.text()
        this_deposit_str = entry_rolling_window.deposit_lineEdit.text()
        balance_now_str = entry_rolling_window.current_balance_lineEdit.text()
        this_rolling_str = entry_rolling_window.this_rolling_lineEdit.text()
        settled_rolling_str = entry_rolling_window.settled_rolling_lineEdit.text()
        rolling_total_str = entry_rolling_window.total_rolling_lineEdit.text()
        win_total_str = entry_rolling_window.total_win_lineEdit.text()

        # 把上面的数据转换成数字格式
        old_balance = round(float(old_balance_str.replace(',', '')), 8)
        this_borrow = round(float(this_borrow_str.replace(',', '')), 8)
        this_return = round(float(this_return_str.replace(',', '')), 8)
        this_win = round(float(this_win_str.replace(',', '')), 8)
        this_cashout = round(float(this_cashout_str.replace(',', '')), 8)
        this_deposit = round(float(this_deposit_str.replace(',', '')), 8)
        balance_now = round(float(balance_now_str.replace(',', '')), 8)
        this_rolling = round(float(this_rolling_str.replace(',', '')), 8)
        settled_rolling = round(float(settled_rolling_str.replace(',', '')), 8)
        rolling_total = round(float(rolling_total_str.replace(',', '')), 8)
        win_total = round(float(win_total_str.replace(',', '')), 8)

        # 不可为空的数据
        player = entry_rolling_window.player_lineEdit.text()
        staff = entry_rolling_window.staff_lineEdit.text()
        rolling_remark = entry_rolling_window.rolling_remark_plainTextEdit.toPlainText()
        win_remark = entry_rolling_window.win_remark_plainTextEdit.toPlainText()

        if not player or not staff or not rolling_remark or not win_remark:
            show_warning(entry_rolling_window, "请填入完整数据！")
            return

        # 写入洗码明细
        # 构造洗码明细表数据
        rolling_detail_index = get_sheet_row(rolling_summary_file_name) - 1
        rolling_detail_data = [rolling_detail_index, entry_date, end_time, acct_name, this_rolling,
                               rolling_total, settled_rolling, staff, rolling_remark]
        # 写入表格
        rolling_detail_result = save_work_book(rolling_detail_data, rolling_summary_file_name)
        if rolling_detail_result:
            show_info(acct_entry_window, "%s已更新保存！" % rolling_detail_result)
        else:
            show_warning(acct_entry_window, "%s无法保存更新，请手动在对应工作表填入数据！" % rolling_detail_result)
            return

        # 写入洗码总账
        # 获取洗码总账，更新该调账目
        rolling_workbook, rolling_sheet = get_workbook(total_rolling_file_name)

        for row in range(6, rolling_sheet.max_row + 1):
            target_cell = rolling_sheet.cell(row=row, column=2)
            target_acct = target_cell.value
            if acct_name == target_acct:
                target_row = target_cell.row
                # 更新洗码记录
                rolling_sheet.cell(row=target_row, column=5, value=rolling_total)
                # 更新总洗码数
                rolling_sheet.cell(row=target_row, column=6, value=rolling_total)
                break

        # 保存总账更新，成功则给出提示
        try:
            rolling_workbook.save(total_rolling_file_name)
            show_info(acct_entry_window, "%s已更新保存！" % total_rolling_file_name.split("/")[-1])
        except:
            show_warning(acct_entry_window, "%s无法保存更新，请手动在对应工作表填入数据！" % total_rolling_file_name.split("/")[-1])
            return

        # 写入上下水明细
        # 构造上下水明细表数据
        win_detail_index = get_sheet_row(win_lose_summary_file_name) - 1
        win_detail_data = [win_detail_index, entry_date, end_time, acct_name, this_win,
                           win_total, staff, win_remark]
        # 写入表格
        win_detail_result = save_work_book(win_detail_data, win_lose_summary_file_name)
        if win_detail_result:
            show_info(acct_entry_window, "%s已更新保存！" % win_detail_result)
        else:
            show_warning(acct_entry_window, "%s无法保存更新，请手动在对应工作表填入数据！" % win_detail_result)
            return

        # 写入上下水总账
        # 获取上下水总账，更新该调账目
        win_workbook, win_sheet = get_workbook(win_lose_total_file_name)

        for row in range(6, win_sheet.max_row + 1):
            target_cell = win_sheet.cell(row=row, column=2)
            target_acct = target_cell.value
            if acct_name == target_acct:
                target_row = target_cell.row
                # 更新洗码记录
                win_sheet.cell(row=target_row, column=5, value=rolling_total)
                # 更新总洗码数
                win_sheet.cell(row=target_row, column=6, value=rolling_total)
                break

        # 保存上下水总账更新，成功则给出提示
        try:
            win_workbook.save(win_lose_total_file_name)
            show_info(acct_entry_window, "%s已更新保存！" % win_lose_total_file_name.split("/")[-1])
        except:
            show_warning(acct_entry_window, "%s无法保存更新，请手动在对应工作表填入数据！" % win_lose_total_file_name.split("/")[-1])
            return

        # 发送报表
        update_message = """
德胜厅客户洗码更新明细    
    
日期：  %s
    
开局时间：  %s  
结束时间：  %s  
    
户名：  %s 
账号：  %s 
    
之前余额：  %s   万
    
本场客户：  %s  
本场出码：  %s   万
本场还码：  %s   万
客户上/下水：  %s   万
取现：  %s   万
本次存/欠款：  %s   万
    
当前账户余额：  %s   万
    
本场洗码：  %s   万
实结洗码：  %s   万
累计洗码：  %s   万
    
跟单员工：  %s 
备注：  %s        
""" % (entry_date, start_time, end_time, acct_name, acct_number, format(old_balance, ','), player,
       format(this_borrow, ','), format(this_return, ','), format(this_win, ','), format(this_cashout, ','),
       format(this_deposit, ','), format(balance_now, ','), format(this_rolling, ','), format(settled_rolling, ','),
       format(rolling_total, ','), staff, rolling_remark)
        send_message(update_message)
        show_info(entry_rolling_window, "洗码信息已发送财务群！")

        # 调用函数关闭洗码录入界面
        close_rolling_entry()

    # 保存现金日记账的函数
    def update_daily():
        # 获取要更新的数据
        entry_date, entry_time = time.strftime("%Y年%m月%d日 %H:%M", time.localtime()).split(' ')

        remark = daily_cash_window.remark_plainTextEdit.toPlainText()
        if not remark:
            show_warning(daily_cash_window, "收/支摘要不能为空！")
            return
        amount_change = daily_cash_window.amount_change_lineEdit.text()
        if not amount_change:
            show_warning(daily_cash_window, "收/支金额不能为空！")
            return

        amount_change = round(float(amount_change))

        # 获取当前年份和月份
        today = datetime.datetime.today()
        year_now = today.year
        month_now = today.month
        # 构造文件名称
        target_file_name = "%d年%d月现金账.xlsx" % (year_now, month_now)

        # 先判断这个文件存在不存在
        file_list = os.listdir(daily_cash_source_path)

        if target_file_name in file_list:
            # 如果当月已经有了现金账表，先获取之前余额
            target_file = daily_cash_source_path + target_file_name
            target_wb = load_workbook(target_file)
            target_ws = target_wb.worksheets[0]
            old_balance = target_ws.cell(row=target_ws.max_row, column=target_ws.max_column).value
            old_balance = round(float(old_balance), 8)

        else:
            # 文件不存在的话，获取最新一月的文件名称
            # 提取文件名里的年份和月份，转成数字格式存入对应表格中
            year_list = []
            month_list = []
            for file in file_list:
                if file.endswith(".xlsx"):
                    year_str, mon_str = re.findall(r"\d+", file)
                    year_list.append(int(year_str))
                    month_list.append(int(mon_str))

            # 获取最新一个月的表格名称
            latest_file = "%d年%d月现金账.xlsx" % (max(year_list), max(month_list))
            # 构建成源文件路径
            source_file = daily_cash_source_path + latest_file

            # 新建一个当前年份和月份的表格，并保存
            wb = Workbook()
            target_file = daily_cash_source_path + target_file_name
            wb.save(target_file)

            # 从最新一月的表格中复制表头，填入新建的表格中
            source_wb = load_workbook(source_file)
            source_sheet = source_wb.worksheets[0]

            target_wb = load_workbook(target_file)
            target_ws = target_wb.worksheets[0]

            for row in range(1, 4):
                for col in range(1, 7):
                    # 提取单元格的源数据
                    source_cell = source_sheet.cell(row=row, column=col)
                    source_value = source_cell.value
                    # 插入到目标单元格
                    target_cell = target_ws.cell(row=row, column=col)
                    target_cell.value = source_value
                    # 复制单元格格式和样式
                    target_cell.data_type = source_cell.data_type
                    target_cell.fill = copy(source_cell.fill)
                    if source_cell.has_style:
                        target_cell._style = copy(source_cell._style)
                        target_cell.font = copy(source_cell.font)
                        target_cell.border = copy(source_cell.border)
                        target_cell.fill = copy(source_cell.fill)
                        target_cell.number_format = copy(source_cell.number_format)
                        target_cell.protection = copy(source_cell.protection)
                        target_cell.alignment = copy(source_cell.alignment)

            # 获取最新余额，填入到新表中的期初余额
            old_balance = source_sheet.cell(row=source_sheet.max_row, column=source_sheet.max_column).value
            target_ws.cell(row=target_ws.max_row, column=target_ws.max_column).value = old_balance

            # 先把新建的表格保存一下
            target_wb.save(target_file)

        # 计算出最新余额
        new_balance = old_balance + amount_change

        # 获取序号
        entry_index = target_ws.max_row - 2

        # 根据录入的金额是正数还是负数，确定要录入的是收入还是支出
        if amount_change > 0:
            data = [entry_index, entry_date, remark, amount_change, "", new_balance]
        else:
            # 构建成一行数据列表
            data = [entry_index, entry_date, remark, "", amount_change, new_balance]

        # 把数据填入表格中
        try:
            save_player_info_book(data, target_file)

            file = target_file.split("/")[-1]
            # 保存成功提示
            show_info(daily_cash_window, "%s表格更新保存成功！" % file)

            # 提示是否关联保存到公司流动资金账和账房PHP现金总账？
            # 1. 直接关联录入
            # 要把金额除以10000，以获取一万为单位的金额录入公司流动资金账和账房PHP现金总账
            # 2. 关闭当前窗口，打开账目录入选择账号界面

            close_daily_cash()
        except:
            # 保存失败提示错误
            show_warning(daily_cash_window, "现金日记账表格更新保存失败！")
            close_daily_cash()

    # 绑定现金日记账录入界面的确认按钮
    daily_cash_window.confirm_Button.clicked.connect(update_daily)

    # 绑定洗码录入界面的保存按钮
    entry_rolling_window.confirm_pushButton.clicked.connect(rolling_update)

    # 绑定账户选择界面的确认按钮
    acct_choose_window.confirmButton.clicked.connect(confirm_acct)

    # 绑定主界面退出按钮
    app_window.exitButton.clicked.connect(app_window.close)

    # 剩下报表按钮的功能实现
    # 1. 考虑实现洗码佣金汇总成报表
    # 2. 考虑实现每月公司流动资金账报表
    # 3. 考虑实现公司每月盈亏账报表
    # 剩下报表按钮的功能实现

    sys.exit(app.exec_())

    # pyinstaller --windowed --onefile --clean --noconfirm main.py


if __name__ == '__main__':
    main()
