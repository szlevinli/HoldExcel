import pytest

@pytest.fixture
def make_wb_and_ws():
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    def _make_wb_and_ws(title):
        ws.title = title
        return wb, ws
    return _make_wb_and_ws