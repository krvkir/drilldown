from pandas import isnull
from matplotlib import cm
from matplotlib.colors import to_hex
from drilldown import Cell


class PageLinkCell(Cell):
    def __init__(self, text, page_name):
        self._text = text
        self._page_name = page_name

    def _get_string(self):
        return self._text

    def _get_link(self):
        return self._page_name


class ColoredCell(Cell):
    """
    >>> str(ColoredCell(10, 100))
    '10.00'

    >>> str(ColoredCell(10, 100, fmt='%.0f'))
    '10'

    >>> from matplotlib import cm
    >>> ColoredCell(10, 100, fmt='%.0f', cmap=cm.RdYlGn).color
    '#d62f27'

    >>> str(ColoredCell(None, 100, fmt='%.0f'))
    ''

    >>> ColoredCell(None, 100, fmt='%.0f').color
    None
    """
    def __init__(self, value, max_value=1.0, cmap=cm.viridis, fmt='%.2f'):
        self._value = value
        self._max_value = max_value
        self._cmap = cmap
        self._fmt = fmt

    def _get_string(self):
        if isnull(self._value):
            return ''
        return self._fmt % self._value

    def _get_color(self):
        if isnull(self._value):
            return
        color_coord = max(0, min(1-1e-9, self._value / self._max_value))
        return to_hex(self._cmap(color_coord))
