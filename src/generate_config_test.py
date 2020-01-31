from generate_config import get_configs

def test_get_configs(mocker, make_wb_and_ws):
    return_value = [
        ('ConfigID', 'FileName', 'SheetName'),
        (1, 'filename1', 'sheetname1'),
        (2, 'filename1', 'sheetname2'),
        (3, 'filename2', 'sheetname1')
    ]
    expect_value = {
        1: {'FileName': 'filename1', 'SheetName': 'sheetname1'},
        2: {'FileName': 'filename1', 'SheetName': 'sheetname2'},
        3: {'FileName': 'filename2', 'SheetName': 'sheetname1'}
    }
    wb, ws = make_wb_and_ws('Config')
    mocker.patch.object(ws, 'iter_rows', return_value=return_value)
    maps = get_configs(wb)
    assert maps == expect_value
