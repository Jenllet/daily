#!/user/bin/python3
# -*- coding: utf-8 -*-
# @Software    : daily
# @Name        : lod2new
# @Authot      : Jabari_Wei
# @Date        : 2023/2/16
# @Comment      :老带新数据统计

import pandas as pd
import xlwings as xw

import achievements
import arrive
import consult
import register

# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':
    # 个人老带新到院
    employee_employee_arrive_old2new = arrive.employee_arrive_old2new().reset_index()
    # 个人老带新建档
    employee_employee_register_old2new = register.employee_register_old2new().reset_index()
    # 个人老带新业绩
    employee_employee_achievements_old2new = achievements.employee_achievements_old2new().reset_index()

    # 老带新业绩月环比
    employee_achievements_old2new_month_tm = achievements.employee_achievements_old2new_month('this_month')
    employee_achievements_old2new_month_lm = achievements.employee_achievements_old2new_month('last_month')

    # 老带新来院月环比
    employee_arrive_zhou_month_tm = arrive.employee_arrive_zhou_month('this_month')
    employee_arrive_zhou_month_lm = arrive.employee_arrive_zhou_month('last_month')

    # 老带新建档月环比
    employee_register_old2new_month_tm = register.employee_register_old2new_month('this_month')
    employee_register_old2new_month_lm = register.employee_register_old2new_month('last_month')

    achievements_merges1 = [employee_achievements_old2new_month_tm, employee_achievements_old2new_month_lm,
                            employee_arrive_zhou_month_tm, employee_arrive_zhou_month_lm, employee_register_old2new_month_tm, employee_register_old2new_month_lm]

    # 个人周业绩统计
    achievements_merge_df1 = pd.concat(achievements_merges1)
    achievements_merge_df1 = achievements_merge_df1.pivot_table(
        index=['客服姓名', '类别'],
        values=['数值', '日期'],
        columns=['日期']
    ).reset_index()

    wb = xw.Book()

    group = wb.sheets['Sheet1']
    group.name = '个人数据月统计'
    group.range('A1').value = achievements_merge_df1

    wb.sheets.add('Sheet2')
    team = wb.sheets('Sheet2')
    team.name = '个人老带新建档'
    team.range('A1').value = employee_employee_register_old2new

    wb.sheets.add('Sheet3')
    team = wb.sheets('Sheet3')
    team.name = '个人老带新到院'
    team.range('A1').value = employee_employee_arrive_old2new

    wb.sheets.add('Sheet4')
    team = wb.sheets('Sheet4')
    team.name = '个人老带新业绩'
    team.range('A1').value = employee_employee_achievements_old2new
