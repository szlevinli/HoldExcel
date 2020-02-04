from generate_config import get_configs
from generate_data import get_data
from generate_formula import get_formulas
from generate_map import get_maps
from tools import get_template
import global_var as GL


def make_template():
    """集成模板数据
    将模板文件中的四个配置项整合到一个字典对象中
    {
        'configs': configs,
        'datum': datum,
        'formulas': formulas,
        'maps': maps
    }

    Returns:
        Dictionary: 集成的模板数据
    """
    try:
        wb = get_template()
        configs = get_configs(wb)
        datum = get_data(wb)
        formulas = get_formulas(wb)
        maps = get_maps(wb)
    finally:
        wb.close()
    return {
        GL.GL_TEMPLATE_KEY_CONFIGS_NAME: configs,
        GL.GL_TEMPLATE_KEY_DATUM_NAME: datum,
        GL.GL_TEMPLATE_KEY_FORMULAS_NAME: formulas,
        GL.GL_TEMPLATE_KEY_MAPS_NAME: maps
    }


if __name__ == "__main__":
    print(f'template: {make_template()}')
