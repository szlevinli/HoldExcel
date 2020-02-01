from generate_template import make_template


def test_make_template(mocker):
    wb = object()
    mock_get_template = mocker.patch(
        'generate_template.get_template',
        return_value=wb)
    mock_get_configs = mocker.patch(
        'generate_template.get_configs', return_value=100)
    mock_get_data = mocker.patch(
        'generate_template.get_data', return_value=200)
    mock_get_formulas = mocker.patch(
        'generate_template.get_formulas', return_value=300)
    mock_get_maps = mocker.patch(
        'generate_template.get_maps', return_value=400)
    t = make_template()
    expect_value = {
        'configs': 100,
        'datum': 200,
        'formulas': 300,
        'maps': 400
    }
    mock_get_template.assert_called_once()
    mock_get_configs.assert_called_once_with(wb)
    mock_get_data.assert_called_once_with(wb)
    mock_get_formulas.assert_called_once_with(wb)
    mock_get_maps.assert_called_once_with(wb)
    assert t == expect_value
