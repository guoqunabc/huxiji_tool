# -*- coding: utf-8 -*-
"""
Created on Sat Aug 15 16:28:00 2020

@author: Madoka
"""
import os
import pyecharts
import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Map, Bar
# from pyecharts.faker import Faker
from pyecharts.globals import ThemeType
from pyecharts.render import make_snapshot
from snapshot_selenium import snapshot


def get_filenames(path):
    files = []
    for path, dir_list, file_list in os.walk(path):
        for file_name in file_list:
            files.append(os.path.join(path, file_name))
    files = [_.split('/')[-1].strip('.xlsx') for _ in files]
    return files


def draw_map(path_root, arg_month, filename):
    path_data = os.path.join(path_root, 'data', arg_month, filename + '.xlsx')
    path_save = os.path.join(path_root, 'result', arg_month, filename + '.html')
    path_png = os.path.join(path_root, 'result', arg_month, filename + '.png')
    data = pd.read_excel(path_data)

    province, sales = list(data["省份"]), list(data["数量"])
    data_list = [list(z) for z in zip(province, sales)]
    sale_max = max(sales)
    series_name = "意向"
    if '中招网' in filename:
        series_name = "销量"

    c = (Map(init_opts=opts.InitOpts(width="1900px", height="910px", theme=ThemeType.ROMANTIC))
         .set_global_opts(
        # title_opts=opts.TitleOpts(title="2020年7月销售意向图"),
        visualmap_opts=opts.VisualMapOpts(
            min_=1,
            max_=sale_max,
            range_text=['数量颜色区间:'],  # 分区间
            # is_piecewise=True,  #定义图例为分段型，默认为连续的图例
            pos_top="middle",  # 分段位置
            pos_left="left",
            orient="vertical",
            split_number=sale_max - 1  # 分成多个区间
        ),
        legend_opts=opts.LegendOpts(is_show=False),
        tooltip_opts=opts.TooltipOpts(is_show=True, is_always_show_content=True),
    ).add(series_name, data_list, maptype="china"))
    c.render(path_save)

    make_snapshot(snapshot, c.render(), path_png, delay=0.5)
    return


def draw_bar(path_root, arg_month, file_list):
    path_save = os.path.join(path_root, 'result', arg_month, 'Bar.html')
    path_png = os.path.join(path_root, 'result', arg_month, 'Bar.png')
    data_0 = pd.read_excel(file_list[0])
    data_1 = pd.read_excel(file_list[1])
    data_2 = pd.read_excel(file_list[2])

    province, sales = list(data_0["省份"]), list(data_0["数量"])
    data_dict_0 = {z[0]: z[1] for z in zip(province, sales)}
    province, sales = list(data_1["省份"]), list(data_1["数量"])
    data_dict_1 = {z[0]: z[1] for z in zip(province, sales)}
    province, sales = list(data_2["省份"]), list(data_2["数量"])
    data_dict_2 = {z[0]: z[1] for z in zip(province, sales)}
    data_list = []
    for k,v in data_dict_0.items():
        tmp_list = [k, data_dict_0.get(k, 0), data_dict_1.get(k, 0), data_dict_2.get(k, 0)]
        data_list.append(tmp_list)

    data_list = sorted(data_list, key=lambda x: x[1], reverse=True)#[:10]
    province = [_[0] for _ in data_list]
    sales_0 = [_[1] for _ in data_list]
    sales_1 = [_[2] for _ in data_list]
    sales_2 = [_[3] for _ in data_list]

    c = (
        Bar(init_opts=opts.InitOpts(width="1900px", height="900px", theme=ThemeType.LIGHT))
            .add_xaxis(xaxis_data=province)
            .add_yaxis(series_name='预期总量', y_axis=sales_0)
            .add_yaxis(series_name='销售意向', y_axis=sales_1)
            .add_yaxis(series_name='招标意向', y_axis=sales_2)
            .set_global_opts(
            xaxis_opts=opts.AxisOpts(name="省份", name_textstyle_opts=opts.TextStyleOpts(font_size = 18)),
            yaxis_opts=opts.AxisOpts(name="数量", axislabel_opts=opts.LabelOpts(font_size = 18)),
            datazoom_opts=opts.DataZoomOpts(range_start=0, range_end=32, start_value=0))
            .set_series_opts(label_opts=opts.LabelOpts(position='top', font_size = 20)
        ))
    c.render(path_save)

    make_snapshot(snapshot, c.render(), path_png, delay=0.5)
    return


# 路径指定
path_root = os.getcwd()
arg_month = '202010'
file_names = get_filenames(os.path.join(path_root, 'data', arg_month))
for _ in file_names:
    draw_map(path_root, arg_month, _)
data_0 = os.path.join(path_root, 'data', arg_month, '预期市场总量.xlsx')
data_1 = os.path.join(path_root, 'data', arg_month, '意向表.xlsx')
data_2 = os.path.join(path_root, 'data', arg_month, '中招网.xlsx')
file_list = [data_0, data_1, data_2]
draw_bar(path_root, arg_month, file_list)
