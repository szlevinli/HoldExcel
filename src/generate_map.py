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
    ws = wb['Map']
    maps = {}
    rows = []
    for row in ws.iter_rows(min_row=ws.min_row, max_row=ws.max_row,
                            min_col=ws.min_column, max_col=ws.max_column,
                            values_only=True):
        rows.append(row)
    for row in rows[1:]:
        tmp1 = maps[row[0]] if row[0] in maps else []
        tmp1.append({'From': row[1], 'To': row[2]})
        maps[row[0]] = tmp1
    return maps


if __name__ == "__main__":
    from tools import get_template
    print(get_maps(get_template()))
