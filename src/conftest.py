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
    return {
        GL.GL_CONFIG_YAML_KEY_TEMPLATE_DIR: 'test_template_dir',
        GL.GL_CONFIG_YAML_KEY_DATA_DIR: 'test_data_dir',
        GL.GL_CONFIG_YAML_KEY_OUT_DIR: 'test_out_dir',
        GL.GL_CONFIG_YAML_KEY_TEMPLATE_FILE_NAME: 'test_template_file_name'
    }
