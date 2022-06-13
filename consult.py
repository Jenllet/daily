#!/user/bin/python3
# -*- coding: utf-8 -*-
# @Software    : daily
# @Name        : consult
# @Authot      : Jabari_Wei
# @Date        : 2022-04-29
# Comment      : 计算咨询量

from pathlib import Path

import numpy as np
import pandas as pd


def base(date):
    if date == 'this_month':
        flag = '截止昨日'
    elif date == 'yesterday':
        flag = '昨日'
    elif date == 'this_week':
        flag = '近7日'
    elif date == 'last_week':
        flag = '上7日'
    elif date == 'last_year':
        flag = '去年同期'
    elif date == '19_year':
        flag = '19年同期'
    else:
        print('date 输入错误。请输入(this_month、yesterday、this_week、last_week、last_year)')

    paths = [
        'E:\\data\\7. other\\下载数据导入\\咨询对话量\\{:s}\\电商平台.xlsx'.format(flag),
        'E:\\data\\7. other\\下载数据导入\\咨询对话量\\{:s}\\社交平台.xlsx'.format(flag),
        'E:\\data\\7. other\\下载数据导入\\咨询对话量\\{:s}\\搜索平台.xlsx'.format(flag),
        'E:\\data\\7. other\\下载数据导入\\咨询对话量\\{:s}\\信息流(集团).xlsx'.format(flag)
    ]

    employee_info_df = pd.read_excel('E:\\data\\5. 资料\\线上客服部名单(2022.5.5).xlsx', sheet_name='Sheet2')
    employee_info_df = employee_info_df[['名单', '渠道', '团队']]
    employee_info_df.rename(columns={'名单': '客服姓名', '渠道': '所属渠道', '团队': '所属组'}, inplace=True)

    def get_excel(path):
        name = Path(path).stem
        df = pd.read_excel(path, header=1)
        df['Unnamed: 2'] = df['Unnamed: 2'].fillna(method='ffill')
        df['组别'] = df['组别'].fillna(method='ffill')
        df = df[['组别', 'Unnamed: 2', '渠道顾问', '对话量(CRM录入)', '咨询量(CRM录入)', '留联量(CRM录入)']]
        df['渠道1'] = name
        df = pd.merge(df, employee_info_df, left_on='渠道顾问', right_on='客服姓名', how='left')
        df['flag'] = df.apply(lambda x: 0 if x['客服姓名'] is np.NaN else 1, axis=1)
        return df

    df0 = get_excel(paths[0])
    df1 = get_excel(paths[1])
    df2 = get_excel(paths[2])
    df3 = get_excel(paths[3])
    dfs = pd.concat([df0, df1, df2, df3])

    return dfs, flag


def group_df(date):
    """
    渠道咨询对话量
    :param date:判定df时间
    :return:df
    """
    dfs, flag = base(date)
    grouped_group_df = dfs.loc[dfs['Unnamed: 2'] == '部门小计（算数求和）：'].fillna(0)
    grouped_group_df = grouped_group_df.groupby('渠道1').sum()
    grouped_group_df = grouped_group_df.reset_index()[['渠道1', '对话量(CRM录入)', '咨询量(CRM录入)', '留联量(CRM录入)']]
    grouped_group_df.rename(columns={'对话量(CRM录入)': '对话量', '咨询量(CRM录入)': '咨询量', '留联量(CRM录入)': '留联量'}, inplace=True)

    df = pd.DataFrame(
        {
            '渠道1': grouped_group_df['渠道1'].tolist() * 3,
            '类别': ['对话量', ] * len(grouped_group_df['渠道1'].tolist()) + ['咨询量', ] * len(
                grouped_group_df['渠道1'].tolist()) + ['留联量', ] * len(grouped_group_df['渠道1'].tolist()),
            '数值': grouped_group_df['对话量'].append(grouped_group_df['咨询量']).append(grouped_group_df['留联量'])
        }
    )
    df = df.groupby(['渠道1', '类别']).sum()
    df['日期'] = flag
    df = df.reset_index()
    print('渠道咨询数据读取成功')
    return df


def team_df(date):
    """
    小组咨询对话量
    :param date:
    :return:
    """
    dfs, flag = base(date)
    grouped_team_df = dfs.loc[dfs['flag'] == 1].groupby('所属组').sum()[['对话量(CRM录入)', '咨询量(CRM录入)', '留联量(CRM录入)']]
    grouped_team_df.rename(columns={'对话量(CRM录入)': '对话量', '咨询量(CRM录入)': '咨询量', '留联量(CRM录入)': '留联量'}, inplace=True)
    grouped_team_df = grouped_team_df.reset_index()

    df = pd.DataFrame(
        {
            '所属组': grouped_team_df['所属组'].tolist() * 3,

            '类别': ['对话量', ] * len(grouped_team_df['所属组'].tolist()) + ['咨询量', ] * len(
                grouped_team_df['所属组'].tolist()) + ['留联量', ] * len(grouped_team_df['所属组'].tolist()),
            '数值': grouped_team_df['对话量'].append(grouped_team_df['咨询量']).append(grouped_team_df['留联量'])
        }
    )

    df = df.groupby(['所属组', '类别']).sum()
    df['日期'] = flag
    df = df.reset_index()
    print('小组咨询数据读取成功')
    return df


def employee_consult(date='this_month'):
    """
    个人咨询量
    :param date:
    :return:
    """
    dfs, flag = base(date)
    df = dfs.loc[dfs['flag'] == 1].groupby('渠道顾问').sum()[
        ['对话量(CRM录入)', '咨询量(CRM录入)', '留联量(CRM录入)']]
    df.rename(columns={'对话量(CRM录入)': '对话量', '咨询量(CRM录入)': '咨询量', '留联量(CRM录入)': '留联量'},
              inplace=True)
    print('个人咨询数据读取成功')
    return df
