import os
import re
import xlrd
import pandas as pd
from tqdm import tqdm

NAMES_MANAGER = ['徐明月', '李想', '杨照', '梁小东', '武思慧', '武真羽', '段玉晶', '王兴龙', '董斌', '贾宇鑫']
ALLOWED_EXT = ['.xls', '.xlsx']


def find_manager_name_in_file(file_name):
    for _name in NAMES_MANAGER:
        if _name in file_name:
            return _name
    raise Exception(f"No manager name in {file_name}")
    return "Error"


def get_manager_names(dir_files):
    managers = []
    for path, dir_list, file_list in os.walk(dir_files):
        for file_name in file_list:
            if os.path.splitext(file_name)[-1] not in ALLOWED_EXT or "~$" in file_name:
                continue
            _name = find_manager_name_in_file(file_name)
            if _name not in managers:
                managers.append(_name)
    return managers


def find_manager_files(name_manager, dir_files):
    assert name_manager in NAMES_MANAGER
    list_file = []
    for path, dir_list, file_list in os.walk(dir_files):
        for file_name in file_list:
            if os.path.splitext(file_name)[-1] not in ALLOWED_EXT:
                continue
            if name_manager in file_name:
                list_file.append(os.path.join(path, file_name))
    return list_file

def get_data(file_path):
    wb = xlrd.open_workbook(filename=file_path)
    sheet1 = wb.sheet_by_index(0)
    n_rows, n_cols = sheet1.nrows, sheet1.ncols
    row_tar, col_tar = 0, 0
    for _ in range(0, n_rows):
        if '按产品分类拜访次数' in sheet1.row_values(_):
            row_tar = _
            break
    for _ in range(0, n_cols):
        if '呼吸机' in str(sheet1.cell_value(row_tar, _)):
            col_tar = _
            break
    cell_tar = sheet1.cell_value(row_tar, col_tar)
    # for symbol in ['：', ':']:
    #     if symbol in cell_tar:
    #         result_list = cell_tar.split(symbol)
    #         break
    try:
        # result = int(result_list[-1])
        result = re.findall(r"\d+", cell_tar)[0]
    except:
        result = 0
    return result


def get_summary(dir_files):
    manager_names = get_manager_names(dir_files)
    df_all_dict = {'姓名': [], '总拜访(填写)': []}
    for manager in manager_names:
        filenames = find_manager_files(manager, dir_files)
        num_total = 0
        for _ in filenames:
            num_huxiji = get_data(_)
            num_total += num_huxiji
        df_dict = {}
        df_dict['姓名'] = manager
        df_dict['总拜访(填写)'] = num_total
        for k, v in df_dict.items():
            df_all_dict[k].append(df_dict[k])
    df_all = pd.DataFrame(df_all_dict)
    return df_all

def get_data_v2(file_path):
    wb = xlrd.open_workbook(filename=file_path)
    sheet1 = wb.sheet_by_index(0)
    n_rows, n_cols = sheet1.nrows, sheet1.ncols
    visit_list = []
    for t_row in range(0, n_rows):
        if '拜访方式' == ''.join(str(sheet1.cell_value(t_row, 0)).split()):
            for t_col in range(1, n_cols):
                if '' != ''.join(str(sheet1.cell_value(t_row, t_col)).split()):
                    temp_pair = [
                        sheet1.cell_value(t_row, t_col),
                        sheet1.cell_value(t_row + 1, t_col),
                        sheet1.cell_value(t_row + 2, t_col)
                    ]
                    visit_list.append(temp_pair)
                    if temp_pair[1] == '' and '拜访' in temp_pair[0]:
                        print(file_path, temp_pair)
    return visit_list

def visit_judge(input_list):
    '''
    根据输入的list，判断本次为哪种形式的拜访
    return: 否定、电话、上门
    '''
    assert len(input_list) == 3
    mark = '否定'
    for _ in ['呼吸机', '高频']:
        if _ in str(input_list[1]) or _ in str(input_list[2]):
            mark = '电话'
            break
    if '电话' == mark:
        for _ in ['上门', '见面', '当面', '科室拜访',
                  '门诊拜访', '面谈', '面访', '办公室拜访', '现场']:
            if _ in str(input_list[0]):
                mark = '上门'
                break
    return mark

def visit_count(input_list):
    '''
    根据输入的list，统计不同方式的拜访次数
    return: {'电话拜访':int,'上门拜访':int}
    '''
    result = {'电话拜访': 0, '上门拜访': 0}
    for _ in input_list:
        mark = visit_judge(_)
        if '电话' == mark:
            result['电话拜访'] += 1
        elif '上门' == mark:
            result['上门拜访'] += 1
        else:
            pass
    return result

def get_summary_v2(dir_files):
    '''
    按照拜访方式统计每个产品经理
    '''
    manager_names = get_manager_names(dir_files)
    df_all_dict = {'姓名': [], '电话拜访': [], '上门拜访': []}
    debug_list = []
    for manager in manager_names:
        filenames = find_manager_files(manager, dir_files)
        pair_list = []
        for _ in filenames:
            temp_list = get_data_v2(_)
            pair_list.extend(temp_list)
        debug_list.extend(pair_list)
        df_dict = visit_count(pair_list)
        df_dict['姓名'] = manager
        for k, v in df_dict.items():
            df_all_dict[k].append(df_dict[k])
    df_all = pd.DataFrame(df_all_dict)
    df_all.loc[:, '总拜访(统计)'] = df_all.loc[:, '电话拜访'] + df_all.loc[:, '上门拜访']
    return df_all


def get_data_v3(file_path):
    wb = xlrd.open_workbook(filename=file_path)
    sheet1 = wb.sheet_by_index(0)
    n_rows, n_cols = sheet1.nrows, sheet1.ncols
    row_tar, col_tar = -1, -1
    for _row in range(0, n_rows):
        if '按产品分类拜访次数' in [str(_).strip() for _ in sheet1.row_values(_row)]:
            row_tar = _row
            break
    if -1 == row_tar: raise Exception(f"{file_path}中未找到按产品分类拜访次数")
    for _col in range(0, n_cols):
        if '呼吸机' in sheet1.cell_value(row_tar, _col):
            col_tar = _col
            break
    if -1 == col_tar: raise Exception(f"{file_path}中未找到按呼吸机拜访次数")
    return




if __name__ == '__main__':
    path_root = os.getcwd()
    path_data = os.path.join(path_root, 'data_zhoubao', '202011')
    path_result = os.path.join(path_root, 'result_zhoubao', '拜访次数统计11月.xlsx')
    # 获取拜访次数统计
    result = get_summary(path_data)
    result.to_excel(path_result)
    # 拜访方式次数统计
    result_way = get_summary_v2(path_data)
    path_result = os.path.join(path_root, 'result_zhoubao', '拜访分类统计11月.xlsx')
    result_way.to_excel(path_result)
    # 表格合并
    assert result.shape[0] == result_way.shape[0]
    result_all = pd.merge(result, result_way, on='姓名')
    columns = ['姓名', '电话拜访', '上门拜访', '总拜访(统计)', '总拜访(填写)']
    result_all = result_all[columns]
    path_result = os.path.join(path_root, 'result_zhoubao', '拜访统计汇总11月.xlsx')
    result_all.to_excel(path_result)
