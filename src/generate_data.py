from tools import make_list


def get_data(wb):
    """数据信息
    数据信息以列表的形式构建，形式如下：
    [
        {'ConfigID': 1, 'ColName1': 'a', ...},
        ...
    ]

    Args:
        wb (Workbook): 配置文件句柄

    Returns:
        Directory: 数据信息
    """
    return make_list(wb, 'Data')


if __name__ == "__main__":
    from tools import get_template
    print(f'configs is {get_data(get_template())}')