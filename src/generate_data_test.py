from generate_data import get_data
import global_var as GL


def test_get_data(mocker, make_wb_and_ws):
    mock_function = mocker.patch('generate_data.make_dict_repeatable')
    wb, _ = make_wb_and_ws(GL.GL_EXECEL_SHEET_DATA_NAME)
    get_data(wb)
    mock_function.assert_called_once_with(wb, GL.GL_EXECEL_SHEET_DATA_NAME)
