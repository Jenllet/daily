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
    # 老带新业绩月环比
    employee_achievements_old2new_month_tm = achievements.employee_achievements_old2new_month('this_month')
    employee_achievements_old2new_month_lm = achievements.employee_achievements_old2new_month('last_month')
    employee_achievements_old2new_month_tmf = achievements.employee_achievements_old2new_month('this_month', 'F')
    employee_achievements_old2new_month_lmf = achievements.employee_achievements_old2new_month('last_month', 'F')

    # 老带新来院月环比
    employee_arrive_zhou_month_tm = arrive.employee_arrive_zhou_month('this_month')
    employee_arrive_zhou_month_lm = arrive.employee_arrive_zhou_month('last_month')
    employee_arrive_zhou_month_tmf = arrive.employee_arrive_zhou_month('this_month', 'F')
    employee_arrive_zhou_month_lmf = arrive.employee_arrive_zhou_month('last_month', 'F')

    # 老带新建档月环比
    employee_register_old2new_month_tm = register.employee_register_old2new_month('this_month')
    employee_register_old2new_month_lm = register.employee_register_old2new_month('last_month')

    # 渠道老带新业绩
    group_achievements_old2new_tm = achievements.group_achievements(date='this_month', j='o2n')
    group_achievements_old2new_lm = achievements.group_achievements(date='last_month', j='o2n')
    group_achievements_old2new_ly = achievements.group_achievements(date='last_year',
                                                                    df=achievements.achievements_df_last_year, j='o2n')

    achievements_merges1 = [employee_achievements_old2new_month_tm, employee_achievements_old2new_month_lm,
                            employee_arrive_zhou_month_tm, employee_arrive_zhou_month_lm,
                            employee_register_old2new_month_tm, employee_register_old2new_month_lm,
                            employee_achievements_old2new_month_tmf, employee_achievements_old2new_month_lmf,
                            employee_arrive_zhou_month_tmf, employee_arrive_zhou_month_lmf]

    group_merges1 = [group_achievements_old2new_tm, group_achievements_old2new_lm, group_achievements_old2new_ly]

    # 个人老带新月数据统计
    achievements_merge_df1 = pd.concat(achievements_merges1)
    achievements_merge_df1 = achievements_merge_df1.pivot_table(
        index=['所属渠道', '所属组', '客服姓名', '类别'],
        values=['数值', '日期'],
        columns=['日期']
    ).reset_index()

    group_merge_df1 = pd.concat(group_merges1)
    group_merge_df1 = group_merge_df1.pivot_table(
        index=['渠道1', '类别'],
        values=['数值', '日期'],
        columns=['日期']
    ).reset_index()

    wb = xw.Book()

    group = wb.sheets['Sheet1']
    group.name = '个人数据月统计'
    group.range('A1').value = achievements_merge_df1

    wb.sheets.add('Sheet2')
    team = wb.sheets('Sheet2')
    team.name = '渠道老带新数据'
    team.range('A1').value = group_merge_df1
