from tools import make_dict_from_3cloumns

def get_maps(wb):
    """获取映射关系
    从模板文件中的 'Map' sheet 工作表中提取映射关系，工作表中的映射关系结构：
    三列，分别表示：ConfigID, From, To
    返回的映射关系为字典类型，结构为：
    {ConfigID: [{'From': 'A', 'To': 'B'}, ...]}
    
    Args:
        wb (openpyxl.Workbook): 模板文件（excel）
    
    Returns:
        Dictionary: 映射关系字典
    """
    return make_dict_from_3cloumns(wb, 'Map')


if __name__ == "__main__":
    from tools import get_template
    print(get_maps(get_template()))
