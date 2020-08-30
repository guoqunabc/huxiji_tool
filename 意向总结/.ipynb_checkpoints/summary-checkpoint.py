import os
import xlrd
import pandas as pd

COLUMNS=['序号','终端医院','合作伙伴','科室','主任','数量','分级',
         '当前进展情况描述','是否库存','器械科长','主管副院长',
         '拜访情况描述','出货价','医院价','已报明年计划','预计招标时间', '经办经理']

def get_filenames(path):
    files = []
    for path,dir_list,file_list in os.walk(path):
        for file_name in file_list:
            files.append(os.path.join(path, file_name))
    return files


def get_name(file_path):
    sep_symbol =['—', '-']
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
        if sheet1.cell_value(_, 1)=='' and sheet1.cell_value(_, 5)=='':
            row_end = _
            break
    df = pd.DataFrame(columns =COLUMNS)
    for _ in range(row_beg, row_end):
        df.loc[df.shape[0]+1] = sheet1.row_values(_)+[manager]
    return df

def get_summary(data_path):
    files = get_filenames(data_path)
    df_list = []
    for _ in files:
        manager = get_name(_)
        df_temp =get_data(_, manager)
        print('经理', manager, ':', df_temp.shape[0])
        df_list.append(df_temp)
    df_all = pd.concat(df_list, axis=0)
    df_all = df_all.reset_index(drop=True)
    return df_all


def compare_table(table_1, table_2):
    '''比较两个表格，标注变化字段，并在同目录生成新的两个表格'''
    df_1=pd.read_excel(table_1)
    df_2 = pd.read_excel(table_2)
    result.to_excel(path_result)
    return


if __name__=='__main__':
    path_root = os.getcwd()
    # path_data = os.path.join(path_root, 'data')
    path_data = os.path.join(path_root, 'data', '202006')
    path_result = os.path.join(path_root, 'result', '结果6月.xlsx')
    result = get_summary(path_data)
    result.to_excel(path_result)
    


