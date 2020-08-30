import os
import xlrd
import numpy as np
import pandas as pd

COLUMNS = ['序号', '终端医院', '合作伙伴', '科室', '主任', '数量', '分级',
           '当前进展情况描述', '是否库存', '器械科长', '主管副院长',
           '拜访情况描述', '出货价', '医院价', '意向把握度', '预计招标时间', '经办经理']


def get_filenames(path):
    files = []
    for path, dir_list, file_list in os.walk(path):
        for file_name in file_list:
            files.append(os.path.join(path, file_name))
    return files


def get_name(file_path):
    sep_symbol = ['—', '-']
    manager_name = ''
    for symbol in sep_symbol:
        if symbol in file_path:
            manager_name = file_path.split(symbol)[-2]
    return manager_name


def get_data(file_path, manager='阿里斯加'):
    # file_path = 'E:\\Workspace\\Tool\\意向总结\\data\\产品意向跟踪表—武真羽—20200414.xlsx'
    wb = xlrd.open_workbook(filename=file_path)
    sheet1 = wb.sheet_by_index(0)
    n_rows, n_cols = sheet1.nrows, sheet1.ncols
    row_beg = 0
    row_end = n_rows
    for _ in range(0, n_rows):
        if 'TwinStream高频叠加喷射手术系统' in sheet1.row_values(_):
            # if 'MostCare血流动力学监测仪' in sheet1.row_values(_):
            row_beg = _+1
            break
    for _ in range(row_beg, n_rows):
        if sheet1.cell_value(_, 1) == '' and sheet1.cell_value(_, 5) == '':
            row_end = _
            break
    df = pd.DataFrame(columns=COLUMNS)
    for _ in range(row_beg, row_end):
        df.loc[df.shape[0]+1] = sheet1.row_values(_)+[manager]
    return df


def get_summary(data_path):
    files = get_filenames(data_path)
    df_list = []
    for _ in files:
        manager = get_name(_)
        df_temp = get_data(_, manager)
        print('经理', manager, ':', df_temp.shape[0])
        df_list.append(df_temp)
    df_all = pd.concat(df_list, axis=0)
    df_all = df_all.reset_index(drop=True)
    return df_all


def compare_helper(item, table_list):
    if item['compare'] not in table_list:
        return '新增项目'
    else:
        return None


def compare_table(table_1, table_2):
    '''比较两个表格，标注变化字段，并在同目录生成新的两个表格'''
    df_1 = pd.read_excel(table_1)
    df_2 = pd.read_excel(table_2)
    df_1['compare'] = df_1.apply(
        lambda row: str(row['终端医院'])+str(row['科室'])+str(row['主任']), axis=1)
    df_2['compare'] = df_2.apply(
        lambda row: str(row['终端医院'])+str(row['科室'])+str(row['主任']), axis=1)

    df_1['减少'] = df_1.apply(
        lambda row: None if row['compare'] in list(df_2['compare']) else '减少', axis=1)
    df_2['新增'] = df_2.apply(
        lambda row: None if row['compare'] in list(df_1['compare']) else '新增', axis=1)

    df_2['变化'] = None
    for _ in df_2[df_2['新增']!='新增']['compare']:
        if df_1[df_1['compare']==_]['当前进展情况描述'].values[0] != df_2[df_2['compare']==_]['当前进展情况描述'].values[0]:
            d_index= df_2[df_2['compare']==_].index
            df_2.loc[d_index, '变化'] = '情况描述变化'

        df_1_bawo, df_2_bawo = -1, -1
        if not np.isnan(list(df_1[df_1['compare'] == _]['意向把握度'])[0]):
            df_1_bawo = list(df_1[df_1['compare'] == _]['意向把握度'])[0]
        if not np.isnan(list(df_2[df_2['compare'] == _]['意向把握度'])[0]):
            df_2_bawo = list(df_2[df_2['compare'] == _]['意向把握度'])[0]
        if df_1_bawo != df_2_bawo:
            d_index = df_2[df_2['compare'] == _].index
            df_2.loc[d_index, '变化'] = '把握度变化'

    df_1 = df_1.drop(['compare'], axis=1)
    df_2 = df_2.drop(['compare'], axis=1)

    path_df_1 = table_1[:-5]+'_对比_上'+'.xlsx'
    path_df_2 = table_2[:-5]+'_对比_下'+'.xlsx'
    df_1.to_excel(path_df_1)
    df_2.to_excel(path_df_2)
    return


if __name__ == '__main__':
    path_root = os.getcwd()
    # path_data = os.path.join(path_root, 'data')
    path_data = os.path.join(path_root, 'data', '202008')
    path_result = os.path.join(path_root, 'result', '意向8月.xlsx')
    result = get_summary(path_data)
    result.to_excel(path_result)

    table_1 = os.path.join(path_root, 'result', '结果7月_new.xlsx')
    table_2 = os.path.join(path_root, 'result', '意向8月.xlsx')
    compare_table(table_1, table_2)
