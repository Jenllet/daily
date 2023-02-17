#!/user/bin/python3
# -*- coding: utf-8 -*-
# @Software    : daily
# @Name        : achievements
# @Authot      : Jabari_Wei
# @Date        : 2022-04-29
# Comment      : 计算业绩
import pandas as pd

import my_global


def fun(x):
    if x['渠道1'] == x['所属渠道']:
        level = 1
    elif (x['渠道1'] == '搜索平台') & (x['所属渠道'] == '信息流(集团)'):
        level = 1
    else:
        level = 0
    return level


cs_achievements_df = pd.read_excel('F:\\data\\7. other\\下载数据导入\\业绩查询.xlsx')

employee_info_df = pd.read_excel('F:\\data\\5. 资料\\线上客服部名单(2022.5.5).xlsx', sheet_name='Sheet2')
employee_info_df = employee_info_df[['名单', '渠道', '团队']]
employee_info_df.rename(columns={'名单': '客服姓名', '渠道': '所属渠道', '团队': '所属组'}, inplace=True)

ana_cs_achievements_df = cs_achievements_df[
    ['客户ID', '结账时间', ' 归属渠道客服', '渠道', '分诊意向一级', '分诊意向二级', '分诊意向三级', '实付', '建档时间']]
ana_cs_achievements_df['月'] = ana_cs_achievements_df['结账时间'].map(lambda x: pd.to_datetime(x).month)
ana_cs_achievements_df['日'] = ana_cs_achievements_df['结账时间'].map(lambda x: pd.to_datetime(x).day)
ana_cs_achievements_df['结账时间'] = ana_cs_achievements_df['结账时间'].map(lambda x: x[0:10])

channel = ana_cs_achievements_df['渠道'].str.split('/', expand=True)
ana_cs_achievements_df['渠道1'] = channel[0]
ana_cs_achievements_df['渠道2'] = channel[1]
ana_cs_achievements_df['渠道3'] = channel[2]
achievements_df = pd.merge(ana_cs_achievements_df, employee_info_df, left_on=' 归属渠道客服', right_on='客服姓名',
                           how='left')
achievements_df['flag'] = achievements_df.apply(fun, axis=1)

# 判断是否是本月建档本月来院
achievements_df['是否本月建档业绩'] = achievements_df.apply(
    lambda x: ((pd.to_datetime(x['建档时间']).month == my_global.this_month) & (pd.to_datetime(
        x['建档时间']).year == my_global.this_year)), axis=1)

# 判断是否是本月
achievements_df['是否本月'] = achievements_df.apply(
    lambda x: pd.to_datetime(x['结账时间']).month == my_global.this_month,
    axis=1)

# 导入去年同期数据(需要更改文件名)
achievements_df_last_year = pd.read_excel('F:\\data\\7. other\\下载数据导入\\去年\\业绩查询.xlsx')
# 导入19年同期数据(需要更改文件名)
achievements_df_19_year = pd.read_excel('F:\\data\\7. other\\下载数据导入\\19年\\业绩查询.xlsx')


def judgement_arrive(date, df):
    """
    判定df时间的方法
    :param date: 判定df时间
    :param df: 对df进行编辑
    :return: 返回时间和df
    """
    if date == 'this_month':
        df = df.loc[df['月'] == my_global.this_month]
        flag = '截止昨日'
    elif date == 'yesterday':
        df = df.loc[df['结账时间'] == my_global.yesterday]
        flag = '昨日'
    elif date == 'this_week':
        df = df.loc[df['结账时间'].isin(my_global.this_week_list)]
        flag = '近7日'
    elif date == 'last_week':
        df = df.loc[df['结账时间'].isin(my_global.last_week_list)]
        flag = '上7日'
    elif date == 'last_month':
        df = df.loc[(df['月'] == my_global.last_month) & (df['日'] <= my_global.this_day)]
        flag = '上月同期'
    elif date == 'last_year':
        df = df[
            ['客户ID', '结账时间', '渠道', '用户组', '分诊意向一级', '分诊意向二级', '分诊意向三级', '实付']
        ]
        df.rename(columns={'用户组': '所属组'}, inplace=True)

        df['结账时间'] = pd.to_datetime(df['结账时间'])
        df['日'] = df['结账时间'].map(lambda x: x.day)

        channel_year = df['渠道'].str.split('/', expand=True)
        df['渠道1'] = channel_year[0]
        df['渠道2'] = channel_year[1]
        df['渠道3'] = channel_year[2]

        df = df.loc[df['日'] <= my_global.this_day]
        df = df.fillna(0)
        df['所属组'] = df['所属组'].map(lambda x: my_global.dic[x])
        flag = '去年同期'
    elif date == '19_year':
        df = df[
            ['客户ID', '结账时间', '渠道', '用户组', '分诊意向一级', '分诊意向二级', '分诊意向三级', '实付']
        ]
        df.rename(columns={'用户组': '所属组'}, inplace=True)

        df['结账时间'] = pd.to_datetime(df['结账时间'])
        df['日'] = df['结账时间'].map(lambda x: x.day)

        channel_year = df['渠道'].str.split('/', expand=True)
        df['渠道1'] = channel_year[0]
        df['渠道2'] = channel_year[1]
        df['渠道3'] = channel_year[2]

        df = df.loc[df['日'] <= my_global.this_day]
        df = df.fillna(0)
        df['所属组'] = df['所属组'].map(lambda x: my_global.dic[x])
        flag = '19年同期'
    else:
        print('date 输入错误。请输入(this_month、yesterday、this_week、last_week)')
    return flag, df


def group_achievements(date, df=achievements_df, j='p'):
    """
    渠道业绩
    :param j:
    :param df:
    :param date: 判定df时间
    :return: 编辑好的df
    """
    flag, df = judgement_arrive(date=date, df=df)
    if j == 'o2n':
        df = df.loc[df['渠道2'] == '老带新']
    if '信息流(集团)' not in list(df['渠道1']):
        df = df.append(
            pd.DataFrame([['信息流(集团)', 0, '新客到院数', flag]], columns=['渠道1', '数值', '类别', '日期']))
    # 得到搜索平台/信息流的数据
    x = df.loc[(df['渠道1'] == '搜索平台') & (df['渠道2'] == '信息流')]
    x = x.groupby('渠道1').sum()['实付'].to_frame()['实付'].values
    if len(x) == 0:
        x = 0

    df = df.groupby('渠道1')['实付'].sum().to_frame()
    df['类别'] = '新客业绩'
    df['日期'] = flag
    df.rename(columns={'实付': '数值'}, inplace=True)
    df = df.reset_index()
    df = df.loc[df['渠道1'].isin(my_global.groups)]
    if '信息流(集团)' not in list(df['渠道1']):
        df = df.append(pd.DataFrame([['信息流(集团)', 0, '渠道业绩', flag]], columns=['渠道1', '数值', '类别', '日期']))

    # 增加信息流数据 减少搜索平台
    df.loc[df['渠道1'] == '信息流(集团)', '数值'] = df.loc[df['渠道1'] == '信息流(集团)']['数值'] + x
    df.loc[df['渠道1'] == '搜索平台', '数值'] = df.loc[df['渠道1'] == '搜索平台']['数值'] - x
    print('渠道业绩数据读取成功')
    return df


def team_achievements(date, df=achievements_df):
    """
    得到小组业绩(按咨询师分)
    :param date:日期控制变量
    :param df:需要编辑的df
    :return:编辑后的df
    """

    flag, df = judgement_arrive(date, df)
    df = df.groupby('所属组').sum()['实付'].to_frame()
    df['类别'] = '新客业绩'
    df['日期'] = flag
    df.rename(columns={'实付': '数值'}, inplace=True)
    df = df.reset_index()
    print('小组业绩数据读取成功')
    return df


team_achievements('this_month')


def team_achievements_for_system(date, df=achievements_df):
    """
    得到小组业绩(按系统小组名分)
    :param date:日期控制变量
    :param df:需要编辑的df
    :return:编辑后的df
    """

    flag, df = judgement_arrive(date, df)
    df['所属组'] = df['所属组'].map(lambda x: my_global.dic[x])
    df = df.groupby('所属组').sum()['实付'].to_frame()
    df['类别'] = '新客业绩'
    df['日期'] = flag
    df.rename(columns={'实付': '数值'}, inplace=True)
    df = df.reset_index()
    print('小组业绩数据读取成功')
    return df


def employee_achievements():
    """
    个人业绩
    :return:
    """
    df = achievements_df.loc[achievements_df['是否本月'] == True]
    df = df.pivot_table(
        index=['所属渠道', '所属组', '客服姓名'],
        values='实付',
        aggfunc={'实付': 'sum'},
        columns='结账时间',
        margins=True,
        margins_name='业绩'
    ).fillna(0)
    df = df.sort_values(by=['结账时间'], axis=1, ascending=False)
    print('个人业绩数据读取成功')
    return df


def employee_achievements_zhou(date, df=achievements_df):
    """
    个人业绩周统计
    :return:
    """
    flag, df = judgement_arrive(date, df)
    df = df.groupby('客服姓名').sum()['实付'].to_frame()
    df['类别'] = '新客业绩'
    df['日期'] = flag
    df.rename(columns={'实付': '数值'}, inplace=True)
    df = df.reset_index()
    print('个人业绩周统计数据读取成功')
    return df


def employee_achievements_old2new_month(date, j="T"):
    """
    个人业绩月统计
    :return:
    """
    flag, df = judgement_arrive(date=date, df=achievements_df)

    if j == "T":
        s = "老带新业绩"
    else:
        s = "本月建档老带新业绩"
        df = df.loc[df['是否本月建档业绩'] == True]
    df = df.loc[df['渠道2'] == '老带新']
    df = df.groupby(['所属渠道', '所属组', '客服姓名']).sum()['实付'].to_frame()
    df['类别'] = s
    df['日期'] = flag
    df.rename(columns={'实付': '数值'}, inplace=True)
    df = df.reset_index()
    print('个人业绩月统计数据读取成功')
    return df
