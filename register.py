#!/user/bin/python3
# -*- coding: utf-8 -*-
# @Software    : daily
# @Name        : register
# @Authot      : Jabari_Wei
# @Date        : 2022-04-29
# Comment      : 计算建档
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
customer_register_df = pd.read_excel('F:\\data\\7. other\\下载数据导入\\渠道客户查询.xlsx')
# 导入员工数据(推荐每周更新一次)
employee_info_df = pd.read_excel('F:\\data\\5. 资料\\线上客服部名单(2022.5.5).xlsx', sheet_name='Sheet2')
employee_info_df = employee_info_df[['名单', '渠道', '团队']]
employee_info_df.rename(columns={'名单': '客服姓名', '渠道': '所属渠道', '团队': '所属组'}, inplace=True)

ana_customer_register_df = customer_register_df[
    ['电话', '建档时间', '渠道客服', '渠道', '一级建档主意向', '二级建档主意向', '三级建档主意向']]
ana_customer_register_df['建档时间'] = ana_customer_register_df['建档时间'].map(lambda x: x[0:10])

channel = ana_customer_register_df['渠道'].str.split('/', expand=True)
ana_customer_register_df['渠道1'] = channel[0]
ana_customer_register_df['渠道2'] = channel[1]
ana_customer_register_df['渠道3'] = channel[2]
register_df = pd.merge(ana_customer_register_df, employee_info_df, left_on='渠道客服', right_on='客服姓名', how='left')
register_df['flag'] = register_df.apply(fun, axis=1)
# 判断是否是本月
register_df['是否本月'] = register_df.apply(lambda x: pd.to_datetime(x['建档时间']).month == my_global.this_month,
                                            axis=1)

# 导入去年同期数据(需要更改文件名)
register_df_last_year = pd.read_excel('F:\\data\\7. other\\下载数据导入\\去年\\渠道客户查询.xlsx')
# 导入19年同期数据(需要更改文件名)
register_df_19_year = pd.read_excel('F:\\data\\7. other\\下载数据导入\\19年\\渠道客户查询.xlsx')


def judgement_arrive(date, df):
    """
    判定df时间的方法
    :param date: 判定df时间
    :param df: 对df进行编辑
    :return: 返回时间和df
    """
    if date == 'this_month':
        df = df.loc[df['是否本月'] == True]
        flag = '截止昨日'
    elif date == 'yesterday':
        df = df.loc[df['建档时间'] == my_global.yesterday]
        flag = '昨日'
    elif date == 'this_week':
        df = df.loc[df['建档时间'].isin(my_global.this_week_list)]
        flag = '近7日'
    elif date == 'last_week':
        df = df.loc[df['建档时间'].isin(my_global.last_week_list)]
        flag = '上7日'
    elif date == 'last_year':
        df = df[
            ['电话', '建档时间', '渠道', '渠道客服组织', '一级建档主意向', '二级建档主意向', '三级建档主意向']
        ]
        df.rename(columns={'渠道客服组织': '所属组'}, inplace=True)
        df['建档时间'] = pd.to_datetime(df['建档时间'])
        df['日'] = df['建档时间'].map(lambda x: x.day)

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
            ['电话', '建档时间', '渠道', '渠道客服组织', '一级建档主意向', '二级建档主意向', '三级建档主意向']
        ]
        df.rename(columns={'渠道客服组织': '所属组'}, inplace=True)
        df['建档时间'] = pd.to_datetime(df['建档时间'])
        df['日'] = df['建档时间'].map(lambda x: x.day)

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


def group_register(date, df=register_df):
    """
    渠道建档数
    :param df:
    :param date: 判定df时间
    :return: 编辑好的df
    """
    flag, df = judgement_arrive(date, df)
    # 得到搜索平台/信息流的数据
    x = df.loc[(df['渠道1'] == '搜索平台') & (df['渠道2'] == '信息流')]
    x = x.groupby('渠道1').count()['电话'].to_frame()['电话'].values
    if len(x) == 0:
        x = 0
    df = df.groupby('渠道1').count()['电话'].to_frame()
    df['类别'] = '建档人数'
    df['日期'] = flag
    df.rename(columns={'电话': '数值'}, inplace=True)
    df = df.reset_index()
    df = df.loc[df['渠道1'].isin(my_global.groups)]
    if '信息流(集团)' not in list(df['渠道1']):
        df = df.append(pd.DataFrame([['信息流(集团)', 0, '建档人数', flag]], columns=['渠道1', '数值', '类别', '日期']))
    # 增加信息流数据 减少搜索平台
    df.loc[df['渠道1'] == '信息流(集团)', '数值'] = df.loc[df['渠道1'] == '信息流(集团)']['数值'] + x
    df.loc[df['渠道1'] == '搜索平台', '数值'] = df.loc[df['渠道1'] == '搜索平台']['数值'] - x
    print('渠道建档数据读取成功')
    return df


def team_register(date, df=register_df):
    """
    得到小组建档数
    :param date:日期控制变量
    :param df:需要编辑的df
    :return:编辑后的df
    """

    flag, df = judgement_arrive(date, df)
    df = df.groupby('所属组').count()['电话'].to_frame()
    df['类别'] = '建档人数'
    df['日期'] = flag
    df.rename(columns={'电话': '数值'}, inplace=True)
    df = df.reset_index()
    print('小组建档数据读取成功')
    return df


def team_register_for_system(date, df=register_df):
    """
    得到小组建档数
    :param date:日期控制变量
    :param df:需要编辑的df
    :return:编辑后的df
    """

    flag, df = judgement_arrive(date, df)
    df = df.groupby('所属组').count()['电话'].to_frame()
    df['类别'] = '建档人数'
    df['日期'] = flag
    df.rename(columns={'电话': '数值'}, inplace=True)
    df = df.reset_index()
    print('小组建档数据读取成功')
    return df


def employee_register():
    """
    个人建档
    :return:
    """
    # df = register_df.loc[register_df['flag'] == 1]
    df = register_df.loc[register_df['是否本月'] == True]
    df = df.pivot_table(
        index=['所属渠道', '所属组', '客服姓名'],
        values='电话',
        aggfunc={'电话': 'count'},
        columns='建档时间',
        margins=True,
        margins_name='建档'
    ).fillna(0)
    df = df.sort_values(by=['建档时间'], axis=1, ascending=False)
    print('个人建档数据读取成功')
    return df


def employee_register_old2new():
    """
    个人老带新建档
    :return:
    """
    df = register_df.loc[(register_df['是否本月'] == True) & (register_df['渠道2'] == '老带新')]
    df = df.pivot_table(
        index=['所属渠道', '所属组', '客服姓名'],
        values='电话',
        aggfunc={'电话': 'count'},
        columns='建档时间',
        margins=True,
        margins_name='老带新建档'
    ).fillna(0)
    df = df.sort_values(by=['建档时间'], axis=1, ascending=False)
    print('个人建档数据读取成功')
    return df
