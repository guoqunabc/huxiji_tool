import os
import re
import xlrd
import pandas as pd
from tqdm import tqdm


def get_manager_names(path):
    managers = []
    for path, dir_list, file_list in os.walk(path):
        for dir_name in dir_list:
            managers.append(os.path.join(path, dir_name))
    return managers


def get_filenames(path):
    files = []
    for path, dir_list, file_list in os.walk(path):
        for file_name in file_list:
            if file_name.endswith('.xlsx'):
                files.append(os.path.join(path, file_name))
    return files


def get_zhoubao_date(file_path):
    zhoubao_date = file_path.split('\\')[-1]
    zhoubao_date = re.findall(r"\d+\.?\d*", zhoubao_date)
    zhoubao_date = zhoubao_date[0].strip('.')
    if len(zhoubao_date) < 8:
        zhoubao_date = '2020' + zhoubao_date
    return zhoubao_date


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
    for symbol in ['：', ':']:
        if symbol in cell_tar:
            result_list = cell_tar.split(symbol)
            break
    try:
        result = int(result_list[-1])
    except:
        result = 0
    return result


def get_summary(data_path):
    manager_names = get_manager_names(data_path)
    df_list = []
    for manager in manager_names:
        filenames = get_filenames(manager)
        manager = manager.split('\\')[-1]
        num_total = 0
        df = pd.DataFrame()
        for _ in filenames:
            num_huxiji = get_data(_)
            # except Exception as e:
            #     print(_, e)
            num_total += num_huxiji
            # zhoubao_date = get_zhoubao_date(_)
            # df[zhoubao_date] = [num_huxiji]
        df['总数'] = [num_total]
        df.rename(index={0: manager}, inplace=True)
        df_list.append(df)
    df_all = pd.concat(df_list, axis=0)
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


def get_summary_v2(data_path):
    '''
    按照拜访方式统计每个产品经理
    '''
    manager_names = get_manager_names(data_path)
    df_all_dict = {'姓名': [], '电话拜访': [], '上门拜访': []}
    debug_list = []
    for manager in manager_names:
        filenames = get_filenames(manager)
        manager = manager.split('\\')[-1]
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
    df_all.loc[:, '总拜访'] = df_all.loc[:, '电话拜访'] + df_all.loc[:, '上门拜访']
    return df_all, debug_list


if __name__ == '__main__':
    path_root = os.getcwd()
    path_data = os.path.join(path_root, 'data_zhoubao', '202008')
    path_result = os.path.join(path_root, 'result_zhoubao', '拜访次数统计8月.xlsx')
    # 获取拜访次数统计
    result = get_summary(path_data)
    print(result)
    result.to_excel(path_result)
    # 拜访方式次数统计
    path_temp = os.path.join(path_data, '段玉晶', '段玉晶周报20200703.xlsx')
    result_way = get_summary_v2(path_data)[0]
    print(result)
    path_result = os.path.join(path_root, 'result_zhoubao', '拜访分类统计8月.xlsx')
    result_way.to_excel(path_result)
    # df_temp = pd.DataFrame(columns = ['拜访方式','拜访内容'])
    # for _ in visit_type:
    #     df_temp.loc[df_temp.shape[0],:] = _
    # path_temp = os.path.join(path_root, 'result_zhoubao', 'debug.xlsx')
    # df_temp.to_excel(path_temp)
    # print(res_temp)
