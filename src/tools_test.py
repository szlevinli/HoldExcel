import pytest
from pathlib import Path
import global_var as GL
from tools import (
    read_config_yaml,
    get_template_file_full_path,
    get_data_file_full_path,
    get_template,
    make_dict_repeatable,
    make_dict_unrepeatable,
    make_list
)


class TestReadConfigYaml:
    def test_normal(self, mocker):
        return_value = {'key', 'value'}
        mock_is_file = mocker.patch('tools.Path.is_file', return_value=True)
        mock_open = mocker.patch('builtins.open')
        mock_load = mocker.patch('tools.yaml.load', return_value=return_value)
        expect_value = return_value
        result_value = read_config_yaml()
        assert result_value == expect_value
        mock_is_file.assert_called_once()
        mock_open.assert_called_once()
        mock_load.assert_called_once()

    def test_except_file_not_found_error(self, mocker):
        mock_is_file = mocker.patch('tools.Path.is_file', return_value=False)
        mock_open = mocker.patch('builtins.open')
        mock_load = mocker.patch('tools.yaml.load')
        with pytest.raises(FileNotFoundError):
            read_config_yaml()
        mock_is_file.assert_called_once()
        assert not mock_open.called
        assert not mock_load.called


def test_get_template_file_full_path(mocker, make_yaml_config):
    return_value = make_yaml_config
    mock_read_config_yaml = mocker.patch('tools.read_config_yaml',
                                         return_value=return_value)
    mock_get_current_work_dir = mocker.patch('tools.get_current_work_dir',
                                             return_value=Path('/test'))
    expect_value = Path('/test/test_template_dir/test_template_file_name')
    result_value = get_template_file_full_path()
    assert result_value == expect_value
    mock_read_config_yaml.assert_called_once()
    mock_get_current_work_dir.assert_called_once()


def test_get_data_file_full_path(mocker, make_yaml_config):
    return_value = make_yaml_config
    mock_read_config_yaml = mocker.patch('tools.read_config_yaml',
                                         return_value=return_value)
    mock_get_current_work_dir = mocker.patch('tools.get_current_work_dir',
                                             return_value=Path('/test'))
    file_name = 'test.xlsx'
    expect_value = Path(f'/test/test_data_dir/{file_name}')
    result_value = get_data_file_full_path(file_name)
    assert result_value == expect_value
    mock_read_config_yaml.assert_called_once()
    mock_get_current_work_dir.assert_called_once()


def test_get_template(mocker):
    mock_load_workbook = mocker.patch('tools.load_workbook')
    mock_get_template_file_full_path = mocker.patch(
        'tools.get_template_file_full_path',
        return_value='/template_file_full_path')
    get_template()
    mock_load_workbook.assert_called_once_with(
        '/template_file_full_path', read_only=True)
    mock_get_template_file_full_path.assert_called_once()


def test_make_dict_repeatable(mocker, make_wb_and_ws):
    return_value = [
        ('id', 'From', 'To'),
        (1, 'A', 'F'),
        (1, 'B', 'D'),
        (1, 'D', 'C'),
        (2, 'E', 'E'),
        (3, 'I', 'I'),
        (2, 'F', 'D')
    ]
    expect_value = {
        1: [
            {'From': 'A', 'To': 'F'},
            {'From': 'B', 'To': 'D'},
            {'From': 'D', 'To': 'C'}
        ],
        2: [
            {'From': 'E', 'To': 'E'},
            {'From': 'F', 'To': 'D'}
        ],
        3: [
            {'From': 'I', 'To': 'I'}
        ]
    }
    sheet_name = 'Temp'
    wb, ws = make_wb_and_ws(sheet_name)
    mocker.patch.object(ws, 'iter_rows', return_value=return_value)
    dicts = make_dict_repeatable(wb, sheet_name)
    assert dicts == expect_value


def test_make_dict_unrepeatable(mocker, make_wb_and_ws):
    return_value = [
        ('ConfigID', 'FileName', 'SheetName'),
        (1, 'filename1', 'sheetname1'),
        (2, 'filename1', 'sheetname2'),
        (3, 'filename2', 'sheetname1'),
        (1, 'repeatK1', 'repeatK2')
    ]
    expect_value = {
        1: {'FileName': 'repeatK1', 'SheetName': 'repeatK2'},
        2: {'FileName': 'filename1', 'SheetName': 'sheetname2'},
        3: {'FileName': 'filename2', 'SheetName': 'sheetname1'}
    }
    sheet_name = 'Temp'
    wb, ws = make_wb_and_ws(sheet_name)
    mocker.patch.object(ws, 'iter_rows', return_value=return_value)
    dicts = make_dict_unrepeatable(wb, sheet_name)
    assert dicts == expect_value


def test_make_list(mocker, make_wb_and_ws):
    return_value = [
        ('ConfigID', 'ColumnName1', 'ColumnName2', 'ColumnName3'),
        ('1', '100', '200', '300'),
        ('1', '1000', '2000', '3000'),
        ('2', '10000', '20000', '30000')
    ]
    expect_value = [
        {
            'ConfigID': '1', 'ColumnName1': '100',
            'ColumnName2': '200', 'ColumnName3': '300'
        },
        {
            'ConfigID': '1', 'ColumnName1': '1000',
            'ColumnName2': '2000', 'ColumnName3': '3000'
        },
        {
            'ConfigID': '2', 'ColumnName1': '10000',
            'ColumnName2': '20000', 'ColumnName3': '30000'
        },
    ]
    sheet_name = 'Temp'
    wb, ws = make_wb_and_ws(sheet_name)
    mocker.patch.object(ws, 'iter_rows', return_value=return_value)
    lists = make_list(wb, sheet_name)
    assert lists == expect_value
