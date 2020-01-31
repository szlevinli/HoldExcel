from tools import (
    make_dict_repeatable,
    make_dict_unrepeatable,
    make_list
)


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
