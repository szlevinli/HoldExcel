from generate_map import get_maps


def test_get_maps(mocker, make_wb_and_ws):
    return_value = [
        ('id', 'From', 'To'),
        (1, 'A', 'F'),
        (1, 'B', 'D'),
        (1, 'D', 'C'),
        (2, 'E', 'E'),
        (2, 'F', 'D'),
        (3, 'I', 'I')
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
    wb, ws = make_wb_and_ws('Map')
    mocker.patch.object(ws, 'iter_rows', return_value=return_value)
    maps = get_maps(wb)
    assert maps == expect_value
