# 这是一个示例 Python 脚本。

# 按 Shift+F10 执行或将其替换为您的代码。
# 按 双击 Shift 在所有地方搜索类、文件、工具窗口、操作和设置。

import pandas as pd
import xlwings as xw

import achievements
import arrive
import consult
import register


def print_hi(name):
    # 在下面的代码行中使用断点来调试脚本。
    print(f'Hi, {name}')  # 按 Ctrl+F8 切换断点。


# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':
    # 渠道到院
    grouped_group_arrive_n = arrive.group_arrive('this_month')
    grouped_group_arrive_y = arrive.group_arrive('yesterday')
    grouped_group_arrive_tw = arrive.group_arrive('this_week')
    grouped_group_arrive_lw = arrive.group_arrive('last_week')
    grouped_group_arrive_ly = arrive.group_arrive('last_year', arrive.arrive_df_last_year)
    grouped_group_arrive_19 = arrive.group_arrive('19_year', arrive.arrive_df_19_year)
    # 小组到院
    grouped_team_arrive_n = arrive.team_arrive('this_month')
    grouped_team_arrive_y = arrive.team_arrive('yesterday')
    grouped_team_arrive_tw = arrive.team_arrive('this_week')
    grouped_team_arrive_lw = arrive.team_arrive('last_week')
    grouped_team_arrive_ly = arrive.team_arrive('last_year', arrive.arrive_df_last_year)
    grouped_team_arrive_19 = arrive.team_arrive('19_year', arrive.arrive_df_19_year)
    # 个人到院
    grouped_employee_arrive = arrive.employee_arrive().reset_index()

    # 渠道建档
    grouped_group_register_n = register.group_register('this_month')
    grouped_group_register_y = register.group_register('yesterday')
    grouped_group_register_tw = register.group_register('this_week')
    grouped_group_register_lw = register.group_register('last_week')
    grouped_group_register_ly = register.group_register('last_year', register.register_df_last_year)
    grouped_group_register_19 = register.group_register('19_year', register.register_df_19_year)
    # 小组建档
    grouped_team_register_n = register.team_register('this_month')
    grouped_team_register_y = register.team_register('yesterday')
    grouped_team_register_tw = register.team_register('this_week')
    grouped_team_register_lw = register.team_register('last_week')
    grouped_team_register_ly = register.team_register('last_year', register.register_df_last_year)
    grouped_team_register_19 = register.team_register('19_year', register.register_df_19_year)
    # 个人建档
    grouped_employee_register = register.employee_register().reset_index()

    # 渠道业绩
    grouped_group_achievements_n = achievements.group_achievements('this_month')
    grouped_group_achievements_y = achievements.group_achievements('yesterday')
    grouped_group_achievements_tw = achievements.group_achievements('this_week')
    grouped_group_achievements_lw = achievements.group_achievements('last_week')
    grouped_group_achievements_ly = achievements.group_achievements('last_year', achievements.achievements_df_last_year)
    grouped_group_achievements_19 = achievements.group_achievements('19_year', achievements.achievements_df_19_year)
    # 小组业绩
    grouped_team_achievements_n = achievements.team_achievements('this_month')
    grouped_team_achievements_y = achievements.team_achievements('yesterday')
    grouped_team_achievements_tw = achievements.team_achievements('this_week')
    grouped_team_achievements_lw = achievements.team_achievements('last_week')
    grouped_team_achievements_ly = achievements.team_achievements('last_year', achievements.achievements_df_last_year)
    grouped_team_achievements_19 = achievements.team_achievements('19_year', achievements.achievements_df_19_year)
    # 个人业绩周环比
    employee_achievements_zhou_tw = achievements.employee_achievements_zhou('this_week')
    employee_achievements_zhou_lw = achievements.employee_achievements_zhou('last_week')
    # 个人业绩
    grouped_employee_achievements = achievements.employee_achievements().reset_index()

    # 渠道咨询
    grouped_group_consult_n = consult.group_df('this_month')
    grouped_group_consult_y = consult.group_df('yesterday')
    grouped_group_consult_tw = consult.group_df('this_week')
    grouped_group_consult_lw = consult.group_df('last_week')
    grouped_group_consult_ly = consult.group_df('last_year')
    grouped_group_consult_19 = consult.group_df('19_year')
    # 小组咨询
    grouped_team_consult_n = consult.team_df('this_month')
    grouped_team_consult_y = consult.team_df('yesterday')
    grouped_team_consult_tw = consult.team_df('this_week')
    grouped_team_consult_lw = consult.team_df('last_week')
    grouped_team_consult_ly = consult.team_df('last_year')
    grouped_team_consult_19 = consult.team_df('19_year')
    # 个人咨询
    grouped_employee_consult = consult.employee_consult().reset_index()

    group_merges1 = [grouped_group_arrive_n, grouped_group_arrive_y, grouped_group_arrive_tw, grouped_group_arrive_lw,
                     grouped_group_arrive_ly, grouped_group_arrive_19,
                     grouped_group_register_n, grouped_group_register_y, grouped_group_register_tw,
                     grouped_group_register_lw, grouped_group_register_ly, grouped_group_register_19,
                     grouped_group_achievements_n, grouped_group_achievements_y, grouped_group_achievements_tw,
                     grouped_group_achievements_lw, grouped_group_achievements_ly, grouped_group_achievements_19,
                     grouped_group_consult_n, grouped_group_consult_y, grouped_group_consult_tw,
                     grouped_group_consult_lw, grouped_group_consult_ly, grouped_group_consult_19]

    team_merges1 = [grouped_team_arrive_n, grouped_team_arrive_y, grouped_team_arrive_tw, grouped_team_arrive_lw,
                    grouped_team_arrive_ly, grouped_team_arrive_19,
                    grouped_team_register_n, grouped_team_register_y, grouped_team_register_tw,
                    grouped_team_register_lw, grouped_team_register_ly, grouped_team_register_19,
                    grouped_team_achievements_n, grouped_team_achievements_y, grouped_team_achievements_tw,
                    grouped_team_achievements_lw, grouped_team_achievements_ly, grouped_team_achievements_19,
                    grouped_team_consult_n, grouped_team_consult_y, grouped_team_consult_tw, grouped_team_consult_lw,
                    grouped_team_consult_ly, grouped_team_consult_19]

    achievements_merges1 = [employee_achievements_zhou_tw, employee_achievements_zhou_lw]

    # 渠道统计
    group_merge_df1 = pd.concat(group_merges1)
    group_merge_df1 = group_merge_df1.pivot_table(
        index=['渠道1', '类别'],
        values=['数值', '日期'],
        columns=['日期']
    ).reset_index()

    # 小组统计
    team_merge_df1 = pd.concat(team_merges1)
    team_merge_df1 = team_merge_df1.pivot_table(
        index=['所属组', '类别'],
        values=['数值', '日期'],
        columns=['日期']
    ).reset_index()

    # 个人周业绩统计
    achievements_merge_df1 = pd.concat(achievements_merges1)
    achievements_merge_df1 = achievements_merge_df1.pivot_table(
        index=['客服姓名', '类别'],
        values=['数值', '日期'],
        columns=['日期']
    ).reset_index()

    # 嘭嘭针统计
    pengpeng = arrive.employee_arrive_pengpeng()

    wb = xw.Book()

    group = wb.sheets['Sheet1']
    group.name = '个人数据周统计'
    group.range('A1').value = achievements_merge_df1

    wb.sheets.add('Sheet2')
    team = wb.sheets('Sheet2')
    team.name = '渠道统计'
    team.range('A1').value = group_merge_df1

    wb.sheets.add('Sheet3')
    team = wb.sheets('Sheet3')
    team.name = '小组统计'
    team.range('A1').value = team_merge_df1

    wb.sheets.add('Sheet4')
    team = wb.sheets('Sheet4')
    team.name = '个人到院'
    team.range('A1').value = grouped_employee_arrive

    wb.sheets.add('Sheet5')
    team = wb.sheets('Sheet5')
    team.name = '个人建档'
    team.range('A1').value = grouped_employee_register

    wb.sheets.add('Sheet6')
    team = wb.sheets('Sheet6')
    team.name = '个人业绩'
    team.range('A1').value = grouped_employee_achievements

    wb.sheets.add('Sheet7')
    team = wb.sheets('Sheet7')
    team.name = '个人咨询'
    team.range('A1').value = grouped_employee_consult

# 访问 https://www.jetbrains.com/help/pycharm/ 获取 PyCharm 帮助
