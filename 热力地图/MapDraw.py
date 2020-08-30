# -*- coding: utf-8 -*-
"""
Created on Sat Aug 15 16:28:00 2020

@author: Madoka
"""
import os
import pyecharts
import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Map
from pyecharts.globals import ThemeType
from pyecharts.render import make_snapshot
from snapshot_selenium import snapshot


def get_filenames(path):
    files = []
    for path, dir_list, file_list in os.walk(path):
        for file_name in file_list:
            files.append(os.path.join(path, file_name))
    files = [_.split('\\')[-1].strip('.xlsx') for _ in files]
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

    c = (
        Map(init_opts=opts.InitOpts(width="2000px", height="900px", theme=ThemeType.ROMANTIC))
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
            legend_opts=opts.LegendOpts(is_show=False)
        )
            .add(series_name, data_list, maptype="china")
    )
    c.render(path_save)

    make_snapshot(snapshot, c.render(), path_png, delay=0.5)
    return


# 路径指定
path_root = os.getcwd()
arg_month = '202008'
file_names = get_filenames(os.path.join(path_root, 'data', arg_month))
for _ in file_names:
    draw_map(path_root, arg_month, _)
