from tools import make_dict_from_3cloumns


def get_formulas(wb):
    """获取单元格公式
    从模板文件中的 'Formula' sheet 工作表中提取公式内容，工作表中的映射关系结构：
    三列，分别表示：ConfigID, ColumnName, Formula
    返回的映射关系为字典类型，结构为：
    {ConfigID: [{'ColumnName': 'A', 'Formula': 'SUM(...)'}, ...]}

    Args:
        wb (openpyxl.Workbook): 模板文件（excel）

    Returns:
        Dictionary: 单元格公式
    """
    return make_dict_from_3cloumns(wb, 'Formula')
