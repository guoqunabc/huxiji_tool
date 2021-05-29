import os
import numpy as np
import pandas as pd


class SummaryWeek:
    def __init__(self):
        self.list_manager = None
        self.dir_file = None
        self.cur_wb = None
        self.data = None

    def set_dir(self, dir_file):
        assert os.path.exists(dir_file), f'文件夹{dir_file}不存在'
        self.dir_file = dir_file

    def set_manager(self, list_manager):
        assert isinstance(list_manager, list), f'{list_manager}不是一个列表'
        self.list_manager = list_manager



if __name__ == '__main__':
    pass