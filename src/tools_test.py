from tools import make_dict_from_3cloumns


def test_make_dict_from_3cloumns(mocker, make_wb_and_ws):
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
    dicts = make_dict_from_3cloumns(wb, sheet_name)
    assert dicts == expect_value
