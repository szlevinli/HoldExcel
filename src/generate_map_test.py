from generate_map import get_maps


def test_get_maps(mocker, make_wb_and_ws):
    mock_function = mocker.patch('generate_map.make_dict_repeatable')
    wb, _ = make_wb_and_ws('Map')
    get_maps(wb)
    mock_function.assert_called_once_with(wb, 'Map')
