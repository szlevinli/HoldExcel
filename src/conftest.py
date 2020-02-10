import pytest
import global_var as GL


@pytest.fixture
def make_wb_and_ws():
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active

    def _make_wb_and_ws(title):
        ws.title = title
        return wb, ws
    return _make_wb_and_ws


@pytest.fixture
def make_yaml_config():
    """模拟返回 YAML 配置文件内容

    Returns:
        Dictionary: YAML 配置文件
    """
    return {
        GL.GL_CONFIG_YAML_KEY_TEMPLATE_DIR: 'test_template_dir',
        GL.GL_CONFIG_YAML_KEY_DATA_DIR: 'test_data_dir',
        GL.GL_CONFIG_YAML_KEY_OUT_DIR: 'test_out_dir',
        GL.GL_CONFIG_YAML_KEY_TEMPLATE_FILE_NAME: 'test_template_file_name'
    }


@pytest.fixture
def mock_data_config():
    """模拟返回 template.xlsx 文件的 Config 数据

    Returns:
        Dictionary: template.xlsx 文件的 Config 数据
    """
    return {
        1: {
            'FileName': '附件5A.xlsx',
            'SheetName': '信息化建设项目年度预算表',
            'InserBeforeRow': 7
        },
        2: {
            'FileName': '附件7A.xlsx',
            'SheetName': '五年预算表',
            'InserBeforeRow': 9
        }
    }


@pytest.fixture
def mock_data_data():
    """模拟返回 template.xlsx 文件的 Data 数据

    Returns:
        Dictionary: template.xlsx 文件的 Data 数据
    """
    return {
        1: [
            {
                '项目名称': '小招通二期',
                '项目内容或需求': '依据国家发改委对双创建设的要求'
            },
            {
                '项目名称': '小招通三期',
                '项目内容或需求': '要求'
            }
        ],
        2: [
            {
                '项目名称': '小招通四期',
                '项目内容或需求': '依据国家发改委对',
                '项目D': '项目D的内容'
            }
        ]
    }


@pytest.fixture
def mock_data_formula():
    """模拟返回 template.xlsx 文件的 formula 数据

    Returns:
        Dictionary: template.xlsx 文件的 formula 数据
    """
    return {
        1: [
            {
                'ColumnName': 'L',
                'Formula': 'I{row}'
            },
            {
                'ColumnName': 'N',
                'Formula': 'SUM(O{row}:T{row})'
            }
        ],
        2: [
            {
                'ColumnName': 'M',
                'Formula': 'Q{row}'
            },
            {
                'ColumnName': 'AB',
                'Formula': 'SUM(AO{row}:AT{row})'
            }
        ]
    }


@pytest.fixture
def mock_data_map():
    """模拟返回 template.xlsx 文件的 map 数据

    Returns:
        Dictionary: template.xlsx 文件的 map 数据
    """
    return {
        1: [
            {'From': '项目名称', 'To': 'B'},
            {'From': '项目内容或需求', 'To': 'C'}
        ],
        2: [
            {'From': '项目名称', 'To': 'B'},
            {'From': '项目内容或需求', 'To': 'C'},
            {'From': '项目D', 'To': 'D'}
        ]
    }


@pytest.fixture
def mock_data_template(mock_data_config, mock_data_data,
                       mock_data_formula, mock_data_map):
    return {
        GL.GL_TEMPLATE_KEY_CONFIGS_NAME: mock_data_config,
        GL.GL_TEMPLATE_KEY_DATUM_NAME: mock_data_data,
        GL.GL_TEMPLATE_KEY_FORMULAS_NAME: mock_data_formula,
        GL.GL_TEMPLATE_KEY_MAPS_NAME: mock_data_map
    }
