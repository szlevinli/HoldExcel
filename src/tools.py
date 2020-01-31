import os
from pathlib import Path
from openpyxl import load_workbook


def get_current_work_dir():
    """当前工作目录

    Returns:
        Path: 当前工作目录
    """
    return Path(os.getcwd())


def get_template_file_full_path():
    """模板文件全路径，包含文件名

    Returns:
        Path: 模板文件全路径，包含文件名
    """
    return get_current_work_dir() / 'template' / 'template.xlsx'


def get_template():
    """返回模板文件

    Returns:
        Workbook: 模板文件
    """
    return load_workbook(get_template_file_full_path(), read_only=True)


def make_dict_repeatable(wb, sheet_name):
    ws = wb[sheet_name]
    dicts = {}
    rows = []
    for row in ws.iter_rows(min_row=ws.min_row, max_row=ws.max_row,
                            min_col=ws.min_column, max_col=ws.max_column,
                            values_only=True):
        rows.append(row)
    keys = rows[0]
    for row in rows[1:]:
        tmp1 = dicts[row[0]] if row[0] in dicts else []
        tmp2 = {}
        for k, v in enumerate(row[1:]):
            tmp2[keys[k+1]] = v
        tmp1.append(tmp2)
        dicts[row[0]] = tmp1
    return dicts


def make_dict_unrepeatable(wb, sheet_name):
    ws = wb[sheet_name]
    rows = []
    for v in ws.iter_rows(min_row=ws.min_row, max_row=ws.max_row,
                          min_col=ws.min_column, max_col=ws.max_column,
                          values_only=True):
        rows.append(v)
    dicts = {}
    # 第一行是配置项名称
    keys = rows[0]
    # 循环配置信息，排除第一行的配置项名称
    for row in rows[1:]:
        tmp = {}
        # 循环配置项，排除第一个配置项
        # 第一个配置项是 configID，将作为 key
        for k, v in enumerate(row[1:]):
            tmp[keys[k+1]] = v
        dicts[row[0]] = tmp
    return dicts


def make_list(wb, sheet_name):
    ws = wb[sheet_name]
    rows = []
    for v in ws.iter_rows(min_row=ws.min_row, max_row=ws.max_row,
                          min_col=ws.min_column, max_col=ws.max_column,
                          values_only=True):
        rows.append(v)
    lists = []
    # 第一行是配置项名称
    keys = rows[0]
    # 循环配置信息，排除第一行的配置项名称
    for row in rows[1:]:
        tmp = {}
        for k, v in enumerate(row):
            tmp[keys[k]] = v
        lists.append(tmp)
    return lists