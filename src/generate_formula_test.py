from generate_formula import get_formulas


def test_get_formulas(mocker, make_wb_and_ws):
    mock_function = mocker.patch('generate_formula.make_dict_repeatable')
    wb, _ = make_wb_and_ws('Formula')
    get_formulas(wb)
    mock_function.assert_called_once_with(wb, 'Formula')
