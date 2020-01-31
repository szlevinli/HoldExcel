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
