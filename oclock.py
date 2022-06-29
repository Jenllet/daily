#!/user/bin/python3
# -*- coding: utf-8 -*-
# @Software    : daily
# @Name        : oclock
# @Authot      : Jabari_Wei
# @Date        : 2022-05-05
# @Comment     : 整点报时

import datetime

import pandas as pd
import xlwings as xw

import my_global


def fun(x):
    if x['渠道1'] == x['所属渠道']:
        level = 1
    elif (x['渠道1'] == '搜索平台') & (x['所属渠道'] == '信息流(集团)'):
        level = 1
    else:
        level = 0
    return level


employee_info_df = pd.read_excel('E:\\data\\5. 资料\\线上客服部名单(2022.5.5).xlsx', sheet_name='Sheet2')
employee_info_df = employee_info_df[['名单', '渠道', '团队']]
employee_info_df.rename(columns={'名单': '客服姓名', '渠道': '所属渠道', '团队': '所属组'}, inplace=True)

oclock_achievements_df = pd.read_excel('E:\\data\\7. other\\下载数据导入\\整点报时\\业绩查询.xlsx')
oclock_achievements_df = oclock_achievements_df[
    ['客户ID', '结账时间', ' 归属渠道客服', '渠道', '用户组', '分诊意向一级', '分诊意向二级', '分诊意向三级', '实付']]
oclock_achievements_df['结账时间'] = oclock_achievements_df['结账时间'].map(lambda x: x[0:10])
# 得到结账时间在本年中的周数
oclock_achievements_df['周'] = oclock_achievements_df['结账时间'].map(
    lambda x: pd.to_datetime(x).week + 1)
oclock_achievements_df['日'] = oclock_achievements_df['结账时间'].map(
    lambda x: pd.to_datetime(x).weekday())

channel = oclock_achievements_df['渠道'].str.split('/', expand=True)
oclock_achievements_df['渠道1'] = channel[0]
oclock_achievements_df['渠道2'] = channel[1]
oclock_achievements_df['渠道3'] = channel[2]

# 判断是否是本月
oclock_achievements_df['是否本月'] = oclock_achievements_df.apply(
    lambda x: pd.to_datetime(x['结账时间']).month == pd.to_datetime(datetime.datetime.now()).month,
    axis=1)

oclock_df = pd.merge(oclock_achievements_df, employee_info_df, left_on=' 归属渠道客服', right_on='客服姓名', how='left')

oclock_df['flag'] = oclock_df.apply(fun, axis=1)


def judgement_arrive(date, df):
    """
    判定df时间的方法
    :param date: 判定df时间
    :param df: 对df进行编辑
    :return: 返回时间和df
    """
    if date == 'this_month':
        df = df.loc[df['是否本月'] == True]
        flag = '本月业绩'
    elif date == 'now':
        df = df.loc[df['结账时间'] == my_global.now]
        flag = '今日业绩'
    elif date == 'now_week':
        df = df.loc[df['周'] == my_global.now_week]
        flag = '本周业绩'
    elif date == 'yes_week':
        df = df.loc[(df['周'] == my_global.now_week - 1) & (df['日'] <= datetime.datetime.now().weekday())]
        flag = '上周业绩'
    else:
        print('date 输入错误。请输入(this_month、now、now_week、yes_week)')
    return flag, df


def group_oclock(date, df=oclock_df):
    """
    整点报时渠道业绩
    :param df:
    :param date: 判定df时间
    :return: 编辑好的df
    """
    flag, df = judgement_arrive(date, df)

    # 得到搜索平台/信息流的数据
    x = df.loc[(df['渠道1'] == '搜索平台') & (df['渠道2'] == '信息流')]
    x = x.groupby('渠道1').sum()['实付'].to_frame()['实付'].values
    if len(x) == 0:
        x = 0

    df = df.groupby('渠道1')['实付'].sum().to_frame()
    df['日期'] = flag
    df.rename(columns={'实付': '数值'}, inplace=True)
    df = df.reset_index()
    df = df.loc[df['渠道1'].isin(my_global.groups)]
    if '信息流(集团)' not in list(df['渠道1']):
        df = df.append(pd.DataFrame([['信息流(集团)', 0, '渠道业绩', flag]], columns=['渠道1', '数值', '类别', '日期']))

    # 增加信息流数据 减少搜索平台
    df.loc[df['渠道1'] == '信息流(集团)', '数值'] = df.loc[df['渠道1'] == '信息流(集团)']['数值'] + x
    df.loc[df['渠道1'] == '搜索平台', '数值'] = df.loc[df['渠道1'] == '搜索平台']['数值'] - x
    df.rename(columns={'渠道1': '所属组'}, inplace=True)
    print('整点业绩数据读取成功')
    return df


def team_oclock(date, df=oclock_df):
    """
    整点小组业绩
    :param date:日期控制变量
    :param df:需要编辑的df
    :return:编辑后的df
    """

    flag, df = judgement_arrive(date, df)

    # 排除信息流组
    df.loc[df.渠道.str.match('团购'), '渠道3'] = '大众'
    df.loc[(df['渠道1'] == '搜索平台') & (df['所属组'] == '信息流组'), 'flag'] = 1
    # 将渠道项中 团购设置为大众
    df.loc[df.渠道.str.contains('团购'), '渠道3'] = '大众'

    df1 = df.groupby('所属组').sum()['实付'].to_frame()
    df1['日期'] = flag
    df1.rename(columns={'实付': '数值'}, inplace=True)
    df1 = df1.reset_index()
    df1 = df1.loc[df1['所属组'].isin(['信息流组', '搜索组', '商城组', '转诊组'])]

    df2 = df.groupby('渠道3').sum()['实付'].to_frame()
    df2['日期'] = flag
    df2.rename(columns={'实付': '数值'}, inplace=True)
    df2 = df2.reset_index()
    df2 = df2.loc[df2.渠道3.str.contains('新氧') | df2.渠道3.str.contains('美呗') | df2.渠道3.str.contains('大众')]
    df2.loc[df2.渠道3.str.contains('新氧'), '渠道3'] = '新氧'
    df2.loc[df2.渠道3.str.contains('美呗'), '渠道3'] = '美呗'
    df2.loc[df2.渠道3.str.contains('大众'), '渠道3'] = '大众'

    if '新氧' not in list(df2['渠道3']):
        df2 = df2.append(pd.DataFrame([['新氧', 0, flag]], columns=['渠道3', '数值', '日期']))
    elif '美呗' not in list(df2['渠道3']):
        df2 = df2.append(pd.DataFrame([['美呗', 0, flag]], columns=['渠道3', '数值', '日期']))
    elif '大众' not in list(df2['渠道3']):
        df2 = df2.append(pd.DataFrame([['大众', 0, flag]], columns=['渠道3', '数值', '日期']))
    df2 = df2.rename(columns={'渠道3': '所属组'})
    df_merge = pd.concat([df1, df2]).reset_index()
    df_merge.pop('index')
    print('整点小组业绩数据读取成功')
    return df_merge


def start():
    """
    运行程序
    :return:
    """
    # 渠道整点
    group_oclock_n = group_oclock('now')
    group_oclock_nw = group_oclock('now_week')
    group_oclock_yw = group_oclock('yes_week')
    group_oclock_tm = group_oclock('this_month')
    # 小组整点
    team_oclock_n = team_oclock('now')
    team_oclock_nw = team_oclock('now_week')
    team_oclock_yw = team_oclock('yes_week')
    team_oclock_tm = team_oclock('this_month')

    oclock_merges = [group_oclock_n, group_oclock_nw, group_oclock_yw, group_oclock_tm, team_oclock_n, team_oclock_nw,
                     team_oclock_yw, team_oclock_tm]
    oclock_merges_df = pd.concat(oclock_merges)
    oclock_merges_df = oclock_merges_df.pivot_table(
        index=['所属组'],
        values=['数值', '日期'],
        columns=['日期'],
        aggfunc='sum'
    ).reset_index()

    wb = xw.Book()
    group = wb.sheets['Sheet1']
    group.name = 'Sheet1'
    group.range('A1').value = oclock_merges_df


# 开始运行
start()
