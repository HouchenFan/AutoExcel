import numpy as np
import os
import sys
import time
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string


# 获取某列的所有值
def get_col_value(ws, column):
    rows = ws.max_row
    column_data = []
    for i in range(1, rows + 1):
        cell_value = ws.cell(row=i, column=column).value
        column_data.append(cell_value)
    return column_data
class Table():
    def __int__(self, wb, str):
        self.ws_raw = wb[str]
        self.ws_grp_rate = wb['最佳组长打分表']
        self.ws_prt_rate = wb['最佳科长打分表']
        self.grp_name_list = []
        for x in range(7, self.ws_grp_rate.max_row):
            cell_value = self.ws_grp_rate.cell(row=x, column=1).value
            self.grp_name_list.append(cell_value)



    def CopyGrpCpltRate(self, str):
        raw_col_list = get_col_value(self.ws_raw, 1)
        grp_name_idx = []
        for name in self.grp_name_list:
            idx = raw_col_list.index(name)
            grp_name_idx.append(idx + 1)
        if len(self.grp_name_list) != len(grp_name_idx):
            print('Error: 小组数量对不上！！！')
            sys.exit(0)
        for i in range(len(grp_name_idx)):
            key_word = 'members joined more than 6 months'
            row_grp_name = grp_name_idx[i]
            row_raw = raw_col_list[row_grp_name:row_grp_name + 100].index(key_word) + row_grp_name + 1
            col_raw = column_index_from_string('AK')
            row_grp = i + 7
            col_grp = column_index_from_string('G')
            cell_value = self.ws_raw.cell(row=row_raw, column=col_raw).value
            self.ws_grp_rate.cell(row=row_grp, column=col_grp).value = cell_value




if __name__ == '__main__':
    new_table_path = r"D:\表格\NewTable.xlsx" #新表格的路径
    quarter = 'Q4'
    #读取表格
    print("开始读取表格！")
    wb_new = load_workbook(new_table_path, data_only=True)
    print("表格读取完成！")

    wb = Table()
    wb.__int__(wb_new, quarter)

    t = time.perf_counter()     #记录当前时间
    wb.CopyGrpCpltRate(quarter)  # 复制已转正员工季度平均完成率（CN)
    print(f'coast:{time.perf_counter() - t:.8f}s') #打印总时间
    wb_new.save(r"D:\表格\NewTable_out.xlsx")
    wb_new.close()



