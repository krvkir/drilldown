import pandas as pd

from .util import Name, NameRange, where, where_col, where_row


def colname(num: int):
    """
    Return a col name (only A-Z), starting from 1.

    Note that in this bizarre numeral system with the base 26:
    - there's no 0 symbol (but let's use it down here for clarity);
    - A is 1;
    - Z is 26 which is A0 (akin to 10 in decimal numeral system);
    - AA is 26+1;
    - AZ is 26+26 which is B0 (akin to 20 in decimals);
    - ZA is 26*26+1 which is A0A.

    >>> colname(1)
    'A'
    >>> colname(26)
    'Z'
    >>> colname(27)
    'AA'
    >>> colname(52)
    'AZ'
    >>> colname(53)
    'BA'
    >>> colname(702)
    'ZZ'
    >>> colname(703)
    'AAA'
    """
    # we assume the user counts from 1 but internally we count from 0
    assert 1 <= num
    name = ""
    while num > 0:
        num, r = divmod(num - 1, 26)
        name = chr(r + ord("A")) + name
    return name


def colnum(name: str):
    assert len(name) == 1
    return ord(name) - ord("A") + 1


def cell(colnum, rownum):
    return f"{colname(colnum)}{rownum}"


class SpreadsheetCell:
    """
    Represents a cell in a spreadsheet.

    >>> cell = SpreadsheetCell(1, 1)
    >>> str(cell)
    'A1'
    >>> cell + (1, 1)
    SpreadsheetCell(B2)
    >>> cell + (26, 0)
    SpreadsheetCell(AA1)
    >>> SpreadsheetCell(2, 2) - (1, 1)
    SpreadsheetCell(A1)
    """

    def __init__(self, col: int, row: int):
        assert col > 0, "Cells can't have negative colunm numbers."
        assert row > 0, "Cells can't have negative row numbers."
        self._col = col
        self._row = row

    def __str__(self):
        return f"{colname(self._col)}{self._row}"

    def __repr__(self):
        return f"{type(self).__name__}({str(self)})"

    def __add__(self, other: tuple[int, int]) -> "SpreadsheetCell":
        cols, rows = other
        return type(self)(col=self._col + cols, row=self._row + rows)

    def __sub__(self, other: tuple[int, int]) -> "SpreadsheetCell":
        cols, rows = other
        return self + (-cols, -rows)

    def fix(self, cols=True, rows=True) -> str:
        return (
            f"{'$' if cols else ''}{colname(self._col)}{'$' if rows else ''}{self._row}"
        )


class SpreadsheetRange:
    """
    Represents a range of cells in a spreadsheet.

    >>> c1 = SpreadsheetCell(1, 1)
    >>> c2 = SpreadsheetCell(3, 3)
    >>> rng = SpreadsheetRange(c1, c2)
    >>> str(rng)
    'A1:C3'
    >>> rng + (1, 1)
    SpreadsheetRange(B2:D4)
    >>> SpreadsheetRange(SpreadsheetCell(2, 2), SpreadsheetCell(4, 4)) - (1, 1)
    SpreadsheetRange(A1:C3)
    """

    def __init__(self, cell0: SpreadsheetCell, cell1: SpreadsheetCell):
        self._cell0 = cell0
        self._cell1 = cell1

    def __str__(self):
        return str(self._cell0) + ":" + str(self._cell1)

    def __repr__(self):
        return f"{type(self).__name__}({str(self)})"

    def fix(
        self, cell0_cols=True, cell0_rows=True, cell1_cols=True, cell1_rows=True
    ) -> str:
        return (
            self._cell0.fix(cell0_cols, cell0_rows)
            + ":"
            + self._cell1.fix(cell1_cols, cell1_rows)
        )

    def __add__(self, other: tuple[int, int]) -> "SpreadsheetRange":
        cell0 = self._cell0 + other
        cell1 = self._cell1 + other
        return type(self)(cell0, cell1)

    def __sub__(self, other: tuple[int, int]) -> "SpreadsheetRange":
        cols, rows = other
        return self + (-cols, -rows)


def where_cell(colname: Name, rowname: Name, df: pd.DataFrame):
    """
    Locate a cell in the dataframe.

    >>> idx = pd.MultiIndex.from_tuples([('r1', 'a'), ('r2', 'b')])
    >>> col = pd.MultiIndex.from_tuples([('c1', 'x'), ('c2', 'y')])
    >>> df = pd.DataFrame([[1, 2], [3, 4]], index=idx, columns=col)
    >>> where_cell(('c1', 'x'), ('r1', 'a'), df)
    SpreadsheetCell(A1)
    >>> where_cell(('c2', 'y'), ('r2', 'b'), df)
    SpreadsheetCell(B2)
    """
    idx_col = 0 if colname is None else where_col(colname, df)[0]
    idx_row = 0 if rowname is None else where_row(rowname, df)[0]
    return SpreadsheetCell(idx_col + 1, idx_row + 1)


def where_range(name: Name | NameRange, df: pd.DataFrame):
    """
    Locate a range in the dataframe.

    >>> idx = pd.MultiIndex.from_tuples([('r1', 'a'), ('r2', 'b')])
    >>> col = pd.MultiIndex.from_tuples([('c1', 'x'), ('c2', 'y')])
    >>> df = pd.DataFrame([[1, 2], [3, 4]], index=idx, columns=col)
    >>> # Assuming we are looking up a range that covers the entire dataframe:
    >>> # where_range(..., df) -> SpreadsheetRange(A1:B2)
    """
    r0, c0, r1, c1 = where(name, df)
    cell0 = SpreadsheetCell(c0 + 1, r0 + 1)
    cell1 = SpreadsheetCell(c1 + 1, r1 + 1)

    return SpreadsheetRange(cell0, cell1)
