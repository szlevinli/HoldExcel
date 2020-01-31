from generate_config import get_configs


def test_get_configs(mocker, make_wb_and_ws):
    mock_function = mocker.patch('generate_config.make_dict_unrepeatable')
    wb, _ = make_wb_and_ws('Config')
    get_configs(wb)
    mock_function.assert_called_once_with(wb, 'Config')
