# -*- coding: utf-8 -*-


def get_row_count(self):
    row_count = self.sheet.nrows
    return row_count
#   获取表格的行数


def get_coulumn_count(self):
    column_count = self.sheet.ncols
    return column_count
#   获取表格列数


def get_merge_cell_value(self,row_index,col_index):
    cell_value = None
    for (min_row, max_row, min_col, max_col) in self.sheet.merged_cells:
        if row_index >= min_row and row_index <= max_row:
            if col_index >= min_col and col_index <= max_col:
                cell_value = self.sheet.cell_value(min_row,min_col)  # 合并单元格的值等于合并第一个单元格的值
                break
            else:
                cell_value = self.sheet.cell_value(row_index, col_index)
        else:
            cell_value = self.sheet.cell_value(row_index, col_index)
    return cell_value
#   获取Excel单元格的数据（包含合并单元格）


def get_all_data(self):
    excel_data_list = []
    row_head = self.sheet.row_values(0,0)
    for row_num in range(1,self.get_row_count()):
        row_dict = {}
        for col_num in range(self.get_col_count()):
            row_dict[row_head[col_num]]=self.ger_merge_cell_value(row_num,col_num)
        excel_data_list.append(row_dict)
    return excel_data_list
    #  把excel数据转换成如下格式的数组
    #       [{"字段名1":"字段值1","字段名2":"字段值2"...},{}]


