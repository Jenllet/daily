#!/user/bin/python3
# -*- coding: utf-8 -*-
# @Software    : daily
# @Name        : arrive
# @Author      : Jabari_Wei
# @Date        : 2022-04-28
# @Comment      : 计算来院
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


# 导入本月数据(不需要更改文件名)
customer_arrive_df = pd.read_excel('F:\\data\\7. other\\下载数据导入\\客户来院查询.xlsx')
# 导入员工数据(推荐每周更新一次)
employee_info_df = pd.read_excel('F:\\data\\5. 资料\\线上客服部名单(2022.5.5).xlsx', sheet_name='Sheet2')
employee_info_df = employee_info_df[['名单', '渠道', '团队']]
employee_info_df.rename(columns={'名单': '客服姓名', '渠道': '所属渠道', '团队': '所属组'}, inplace=True)

ana_customer_arrive_df = customer_arrive_df[
    ['客户ID', '接待时间', ' 归属渠道客服', '渠道', '分诊意向一级', '分诊意向二级', '分诊意向三级', '首次/二次来院']
]
ana_customer_arrive_df['接待时间'] = ana_customer_arrive_df['接待时间'].map(lambda x: x[0:10])
channel = ana_customer_arrive_df['渠道'].str.split('/', expand=True)
ana_customer_arrive_df['渠道1'] = channel[0]
ana_customer_arrive_df['渠道2'] = channel[1]
ana_customer_arrive_df['渠道3'] = channel[2]
arrive_df = pd.merge(ana_customer_arrive_df, employee_info_df, left_on=' 归属渠道客服', right_on='客服姓名', how='left')
# arrive_df = arrive_df.drop_duplicates('客户ID')
arrive_df['flag'] = arrive_df.apply(fun, axis=1)
arrive_df['是否本月'] = arrive_df.apply(lambda x: pd.to_datetime(x['接待时间']).month == my_global.this_month, axis=1)

# 导入去年同期数据(需要更改文件名)
arrive_df_last_year = pd.read_excel('F:\\data\\7. other\\下载数据导入\\去年\\客户来院查询.xlsx')
arrive_df_last_year = arrive_df_last_year.drop_duplicates('客户ID')
# 导入19年同期数据(需要更改文件名)
arrive_df_19_year = pd.read_excel('F:\\data\\7. other\\下载数据导入\\19年\\客户来院查询.xlsx')
arrive_df_19_year = arrive_df_19_year.drop_duplicates('客户ID')


def judgement_arrive(date, df):
    """
    判定df时间的方法
    :param date: 判定df时间
    :param df: 对df进行编辑
    :return: 返回时间和df
    """
    df = df.drop_duplicates('客户ID')
    if date == 'this_month':
        df = df.loc[df['是否本月'] == True]
        flag = '截止昨日'
    elif date == 'yesterday':
        df = df.loc[df['接待时间'] == my_global.yesterday]
        flag = '昨日'
    elif date == 'this_week':
        df = df.loc[df['接待时间'].isin(my_global.this_week_list)]
        flag = '近7日'
    elif date == 'last_week':
        df = df.loc[df['接待时间'].isin(my_global.last_week_list)]
        flag = '上7日'
    elif date == 'last_year':
        df = df[
            ['客户ID', '接待时间', '渠道', '用户组', '分诊意向一级', '分诊意向二级', '分诊意向三级']
        ]
        df.rename(columns={'用户组': '所属组'}, inplace=True)
        df['接待时间'] = pd.to_datetime(df['接待时间'])
        df['日'] = df['接待时间'].map(lambda x: x.day)

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
            ['客户ID', '接待时间', '渠道', '用户组', '分诊意向一级', '分诊意向二级', '分诊意向三级']
        ]
        df.rename(columns={'用户组': '所属组'}, inplace=True)
        df['接待时间'] = pd.to_datetime(df['接待时间'])
        df['日'] = df['接待时间'].map(lambda x: x.day)

        channel_year = df['渠道'].str.split('/', expand=True)
        df['渠道1'] = channel_year[0]
        df['渠道2'] = channel_year[1]
        df['渠道3'] = channel_year[2]

        df = df.loc[df['日'] <= my_global.this_day]
        df = df.fillna(0)
        df['所属组'] = df['所属组'].map(lambda x: my_global.dic[x])
        flag = '19年同期'
    else:
        print('date 输入错误。请输入(this_month、yesterday、this_week、last_week、last_year、19_year)')
    return flag, df


def group_arrive(date, df=arrive_df):
    """
    渠道首次到院数
    :param df:
    :param date: 判定df时间
    :return: 编辑好的df
    """
    df = df.loc[df['首次/二次来院'] == '首次']
    df = df.drop_duplicates('客户ID')
    flag, df = judgement_arrive(date, df)
    # 得到搜索平台/信息流的数据
    x = df.loc[(df['渠道1'] == '搜索平台') & (df['渠道2'] == '信息流')]
    x = x.groupby('渠道1').count()['客户ID'].to_frame()['客户ID'].values
    if len(x) == 0:
        x = 0
    df = df.groupby('渠道1').count()['客户ID'].to_frame()
    df['类别'] = '首次来院人数'
    df['日期'] = flag
    df.rename(columns={'客户ID': '数值'}, inplace=True)
    df = df.reset_index()
    df = df.loc[df['渠道1'].isin(my_global.groups)]
    if '信息流(集团)' not in list(df['渠道1']):
        df = df.append(
            pd.DataFrame([['信息流(集团)', 0, '新客到院数', flag]], columns=['渠道1', '数值', '类别', '日期']))
    # 增加信息流数据 减少搜索平台
    df.loc[df['渠道1'] == '信息流(集团)', '数值'] = df.loc[df['渠道1'] == '信息流(集团)']['数值'] + x
    df.loc[df['渠道1'] == '搜索平台', '数值'] = df.loc[df['渠道1'] == '搜索平台']['数值'] - x
    print('渠道到院数据读取成功')
    return df


def group_arrive_total(date, df=arrive_df):
    """
    渠道到院数
    :param df:
    :param date: 判定df时间
    :return: 编辑好的df
    """
    df = df.drop_duplicates('客户ID')
    flag, df = judgement_arrive(date, df)
    # 得到搜索平台/信息流的数据
    x = df.loc[(df['渠道1'] == '搜索平台') & (df['渠道2'] == '信息流')]
    x = x.groupby('渠道1').count()['客户ID'].to_frame()['客户ID'].values
    if len(x) == 0:
        x = 0
    df = df.groupby('渠道1').count()['客户ID'].to_frame()
    df['类别'] = '来院人数'
    df['日期'] = flag
    df.rename(columns={'客户ID': '数值'}, inplace=True)
    df = df.reset_index()
    df = df.loc[df['渠道1'].isin(my_global.groups)]
    if '信息流(集团)' not in list(df['渠道1']):
        df = df.append(
            pd.DataFrame([['信息流(集团)', 0, '新客到院数', flag]], columns=['渠道1', '数值', '类别', '日期']))
    # 增加信息流数据 减少搜索平台
    df.loc[df['渠道1'] == '信息流(集团)', '数值'] = df.loc[df['渠道1'] == '信息流(集团)']['数值'] + x
    df.loc[df['渠道1'] == '搜索平台', '数值'] = df.loc[df['渠道1'] == '搜索平台']['数值'] - x
    print('渠道到院数据读取成功')
    return df


def team_arrive(date, df=arrive_df):
    """
    得到小组首次到院数
    :param date:日期控制变量
    :param df:需要编辑的df
    :return:编辑后的df
    """
    df = df.loc[df['首次/二次来院'] == '首次']
    df = df.drop_duplicates('客户ID')
    flag, df = judgement_arrive(date, df)
    df = df.groupby('所属组').count()['客户ID'].to_frame()
    df['类别'] = '首次来院人数'
    df['日期'] = flag
    df.rename(columns={'客户ID': '数值'}, inplace=True)
    df = df.reset_index()
    print('小组到院数据读取成功')
    return df


def team_arrive_total(date, df=arrive_df):
    """
    得到小组到院数
    :param date:日期控制变量
    :param df:需要编辑的df
    :return:编辑后的df
    """
    df = df.drop_duplicates('客户ID')
    flag, df = judgement_arrive(date, df)
    df = df.groupby('所属组').count()['客户ID'].to_frame()
    df['类别'] = '来院人数'
    df['日期'] = flag
    df.rename(columns={'客户ID': '数值'}, inplace=True)
    df = df.reset_index()
    print('小组到院数据读取成功')
    return df


# 个人到院
def employee_arrive(df=arrive_df):
    """
    个人首次到院
    :return:
    """
    df = df.loc[df['首次/二次来院'] == '首次']
    df = df.drop_duplicates('客户ID')
    df = df.loc[df['是否本月'] == True]
    df = df.pivot_table(
        index=['所属渠道', '所属组', '客服姓名'],
        values='客户ID',
        aggfunc=lambda x: len(x.unique()),
        columns='接待时间',
        margins=True,
        margins_name='到院'
    ).fillna(0)
    df = df.sort_values(by=['接待时间'], axis=1, ascending=False)
    print('个人到院数据读取成功')
    return df


def employee_arrive_zhou(date, df=arrive_df):
    """
    个人首次到院周统计
    :return:
    """
    df = df.loc[df['首次/二次来院'] == '首次']
    df = df.drop_duplicates('客户ID')
    flag, df = judgement_arrive(date, df)
    df = df.groupby('客服姓名').count()['客户ID'].to_frame()
    df['类别'] = '首次来院'
    df['日期'] = flag
    df.rename(columns={'客户ID': '数值'}, inplace=True)
    df = df.reset_index()
    print('个人首次来院周统计数据读取成功')
    return df


def employee_arrive_old2new():
    """
    个人老带新首次来院
    :return:
    """
    df = arrive_df.loc[arrive_df['首次/二次来院'] == '首次']
    df = arrive_df.loc[(arrive_df['是否本月'] == True) & (arrive_df['渠道2'] == '老带新')]
    df = df.pivot_table(
        index=['所属渠道', '所属组', '客服姓名'],
        values='客户ID',
        aggfunc={'客户ID': 'count'},
        columns='接待时间',
        margins=True,
        margins_name='老带新到院'
    ).fillna(0)
    df = df.sort_values(by=['接待时间'], axis=1, ascending=False)
    print('个人老带新首次到院数据读取成功')
    return df
