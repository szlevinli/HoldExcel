from tools import make_dict_unrepeatable
import global_var as GL


def get_configs(wb):
    """配置信息
    配置信息以字典的形式构建，形式如下：
    {configID: {key: value, ...}}

    Args:
        wb (Workbook): 配置文件句柄

    Returns:
        Directory: 配置信息
    """
    return make_dict_unrepeatable(wb, GL.GL_EXECEL_SHEET_CONFIG_NAME)


if __name__ == "__main__":
    from tools import get_template
    print(f'configs is {get_configs(get_template())}')
