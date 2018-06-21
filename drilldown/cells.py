from pandas import isnull, notnull
from drilldown import Cell


class PageLinkCell(Cell):
    def __init__(self, text, page_name):
        self._text = text
        self._page_name = page_name

    def _get_string(self):
        return self._text

    def _get_link(self):
        return self._page_name
