#!/user/bin/python3
# -*- coding: utf-8 -*-
# @Software    : daily
# @Name        : global
# @Authot      : Jabari_Wei
# @Date        : 2022-04-28
import datetime

import pandas as pd

# 全局变量
groups = ['信息流(集团)', '搜索平台', '电商平台', '社交平台']

# 昨天所在的月份
this_month = pd.to_datetime(datetime.datetime.now() - datetime.timedelta(days=1)).month

# 昨天所在的日份
this_day = pd.to_datetime(datetime.datetime.now() - datetime.timedelta(days=1)).day

# 得到昨天的日期
yesterday = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")

# 得到7天前的日期
this_week_day = (datetime.datetime.now() - datetime.timedelta(days=7)).strftime("%Y-%m-%d")

# 得到14天前的日期
last_week_day = (datetime.datetime.now() - datetime.timedelta(days=14)).strftime("%Y-%m-%d")

# 设置近7日时间序列
this_week_list = pd.date_range(this_week_day, periods=7).strftime("%Y-%m-%d")

# 设置上7日时间序列
last_week_list = pd.date_range(last_week_day, periods=7).strftime("%Y-%m-%d")

# 设置当前时间在今年的周数
now_week = pd.to_datetime(datetime.datetime.now()).week + 1

# 设置当前时间在今年的周数
now = datetime.datetime.now().strftime("%Y-%m-%d")

# 重点项目
key_program = ['年轻化面部紧致提升仪器紧肤', '美容及修复鼻部', '年轻化面部祛斑', '年轻化眼部', '美容及修复眼部', '美容及修复胸部', '年轻化面部填充注射填充']

dic = {
    '客服A组': '客服A组',
    '客服C组': '客服C组',
    '整形组（刘芮岑组）': '整形组',
    '整形二组': '整形组',
    '皮肤组（徐颖组）': '整形组',
    '整形组': '整形组',
    '搜索客服': '整形组',
    '信息流组（林玲组）': '信息流组',
    '信息流客服': '信息流组',
    '信息流组': '信息流组',
    '电商客服': '商城',
    '客服B组': '商城',
    '电商组': '商城',
    '商城': '商城',
    '新媒体咨询组': '订阅号',
    '订阅号': '订阅号',
    '新媒体拓客组': '拓客',
    '拓客': '拓客',
    '拓客二组': '拓客',
    '拓客一组': '拓客',
    '新媒体咨询': 'IP',
    'IP客服': 'IP',
    'IP': 'IP',
    '新媒体': '其他',
    '美容牙科': '其他',
    '美容皮肤科': '其他',
    '人力资源部': '其他',
    '现场咨询二组': '其他',
    '回访中心': '其他',
    '口腔事业部': '其他',
    '信息流咨询组': '其他',
    '二级咨询点': '其他',
    '信息管理': '其他',
    '市场营销部': '其他',
    '信息部': '其他',
    '现场咨询一组': '其他',
    '现场咨询三组': '其他',
    '现场咨询四组': '其他',
    '专家助理': '其他',
    '集团信息部': '其他',
    '线上客服部': '其他',
    '社交': '其他',
    '其他': '其他',
    0: '其他',
    '普通病区': '其他',
    '美容外科': '其他',
    '麻醉科': '其他',
    '美容皮肤科-脱毛': '其他',
    '皮肤导诊': '其他',
    '维养中心': '其他'
}
# 对去年的组进行转换
