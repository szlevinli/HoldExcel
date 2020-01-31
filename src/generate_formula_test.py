from generate_formula import get_formulas


def test_get_formulas(mocker, make_wb_and_ws):
    return_value = [
        ('id', 'ColumnName', 'Formula'),
        (1, 'A', 'A'),
        (1, 'B', 'SUM(O{row}:T{row})'),
        (2, 'E', 'E'),
        (2, 'F', 'D'),
        (3, 'I', 'I')
    ]
    expect_value = {
        1: [
            {'ColumnName': 'A', 'Formula': 'A'},
            {'ColumnName': 'B', 'Formula': 'SUM(O{row}:T{row})'}
        ],
        2: [
            {'ColumnName': 'E', 'Formula': 'E'},
            {'ColumnName': 'F', 'Formula': 'D'}
        ],
        3: [
            {'ColumnName': 'I', 'Formula': 'I'}
        ]
    }
    wb, ws = make_wb_and_ws('Formula')
    mocker.patch.object(ws, 'iter_rows', return_value=return_value)
    maps = get_formulas(wb)
    assert maps == expect_value
