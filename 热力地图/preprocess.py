import os
import xlrd
import pandas as pd
from functools import reduce


def read_data_xlsx_v1(file_path, sheet_name):
    '''
    处理带不规则表头的xlsx数据
    '''
    workbook = xlrd.open_workbook(filename=file_path)
    sheet = workbook.sheet_by_name(sheet_name)
    n_rows, n_cols = sheet.nrows, sheet.ncols
    row_beg = 0
    row_end = n_rows
    for _ in range(0, n_rows):  # 判断表头开始的行
        if reduce(lambda x, y: x + y, ['' == _ for _ in sheet.row_values(_)]) < 3:
            row_beg = _
            break
    col_name = sheet.row_values(row_beg)
    col_prov = col_name.index('地区')

    dict_prov = {}
    for _ in range(row_beg + 1, n_rows):
        temp_prov = sheet.cell_value(_, col_prov)
        if temp_prov in dict_prov:
            dict_prov[temp_prov] += 1
        else:
            dict_prov[temp_prov] = 1

    list_prov, list_num = [], []
    for k, v in dict_prov.items():
        list_prov.append(k.strip('省'))
        list_num.append(v)
    # list_prov = [_.strip('省') for _ in list_prov]

    df = pd.DataFrame({'省份': list_prov, '数量': list_num})
    return df


# 预处理并存取pandas格式省市统计数据
path_root = os.getcwd()
arg_month = '202008'
path_data = os.path.join(path_root, 'data_raw', arg_month, '2020年中招网信息整理 20200828.xlsx')
data_df = read_data_xlsx_v1(path_data, '高频呼吸机')

path_save = os.path.join(path_root, 'data', arg_month, '2020年中招网信息.xlsx')
data_df.to_excel(path_save, index=False)

# 预期市场总量
path_data_1 = os.path.join(path_root, 'data', arg_month, '2020年中招网信息.xlsx')
path_data_2 = os.path.join(path_root, 'data_raw', arg_month, '8月意向表.xlsx')

df_1 = pd.read_excel(path_data_1)
df_2 = pd.read_excel(path_data_2)
df_all = pd.merge(df_1, df_2, how='outer', on=['省份'])
df_all = df_all.fillna(0)
df_all.loc[:, '数量'] = df_all['数量_x'] + df_all['数量_y']
df_all = df_all[['省份', '数量']]

path_save = os.path.join(path_root, 'data', arg_month, '预期市场总量.xlsx')
df_all.to_excel(path_save, index=False)
