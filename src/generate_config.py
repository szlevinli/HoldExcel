def get_configs(wb):
    """配置信息
    配置信息以字典的形式构建，形式如下：
    {configID: {key: value, ...}}
    
    Args:
        wb (Workbook): 配置文件句柄
    
    Returns:
        Directory: 配置信息
    """
    ws = wb['Config']
    rows = []
    for v in ws.iter_rows(min_row=ws.min_row, max_row=ws.max_row,
                          min_col=ws.min_column, max_col=ws.max_column,
                          values_only=True):
        rows.append(v)
    configs = {}
    # 第一行是配置项名称
    keys = rows[0]
    # 循环配置信息，排除第一行的配置项名称
    for row in rows[1:]:
        tmp = {}
        # 循环配置项，排除第一个配置项
        # 第一个配置项是 configID，将作为 key
        for k, v in enumerate(row[1:]):
            tmp[keys[k+1]] = v
        configs[row[0]] = tmp
    return configs


if __name__ == "__main__":
    from tools import get_template
    print(f'configs is {get_configs(get_template())}')
