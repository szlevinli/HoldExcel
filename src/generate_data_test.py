from generate_data import get_data


def test_get_data(mocker, make_wb_and_ws):
    mock_function = mocker.patch('generate_data.make_list')
    wb, _ = make_wb_and_ws('Data')
    get_data(wb)
    mock_function.assert_called_once_with(wb, 'Data')
