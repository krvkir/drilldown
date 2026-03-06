from typing import Any, Optional

import pandas as pd

"""
Helpers for finding locations of dataframe cells in a spreadsheet before
those spreadsheets are formed.

Useful for building formulas or specifying locations for conditional
formatting.
"""

# ... we can test with `isinstance` only for the first of the types.
ScalarName = str | int
# ... for the rest of them, we test against generics used in them, e.g. a list or a tuple
CompoundName = tuple[ScalarName]
CompoundNameRange = list[CompoundName]
Name = ScalarName | CompoundName
NameRange = list[Name]
SpreadsheetCoords = tuple[int, int]
SpreadsheetRange1d = tuple[int, int]
SpreadsheetRange2d = tuple[int, int, int, int]


def _ensure_tuple(name: Name) -> CompoundName:
    """Make sure that name is wrapped in tuple if it's scalar."""
    if not isinstance(name, tuple):
        name = (name,)
    return name


def _fetch_names(name: Name | NameRange) -> CompoundNameRange:
    """
    Name may be a scalar, a tuple (compound name) or a list of two names.
    Process those options and return two tuple names.
    """
    if isinstance(name, list):
        assert len(name) >= 2
        name1, name2 = name
    else:
        name1 = name
        name2 = name

    name1 = _ensure_tuple(name1)
    name2 = _ensure_tuple(name2)
    assert len(name1) == len(name2)

    return [name1, name2]


def where_col(
    name: Name | NameRange, df: pd.DataFrame, skip_cols: int = 0
) -> Optional[SpreadsheetRange1d]:
    """
    Find column(s) with the given (scalar or compound) name.

    >>> import pandas as pd
    >>> index = pd.MultiIndex.from_tuples([('R1', 'r1'), ('R1', 'r2')])
    >>> columns = pd.MultiIndex.from_tuples([('C1', 'c1'), ('C1', 'c2')])
    >>> df = pd.DataFrame([[1, 2], [3, 4]], index=index, columns=columns)
    >>> where_col(('C1', 'c1'), df)
    (2, 2)
    >>> where_col([('C1', 'c1'), ('C1', 'c2')], df)
    (2, 3)
    """
    skip_cols += df.index.nlevels - 1
    name1, name2 = _fetch_names(name)
    assert df.columns.nlevels >= len(name1)
    cols = list(
        zip(*[df.columns.get_level_values(-i) for i in range(1, len(name1) + 1)][::-1])
    )
    if name1 in cols and name2 in cols:
        idx1 = cols.index(name1) + 1
        idx2 = cols.index(name2) + 1
        return (skip_cols + idx1, skip_cols + idx2)


def where_row(
    name: Name | NameRange, df: pd.DataFrame, skip_rows: int = 3
) -> Optional[SpreadsheetRange1d]:
    """
    Find row(s) with the given (scalar or compound) name.

    >>> import pandas as pd
    >>> index = pd.MultiIndex.from_tuples([('R1', 'r1'), ('R1', 'r2')])
    >>> columns = pd.MultiIndex.from_tuples([('C1', 'c1'), ('C1', 'c2')])
    >>> df = pd.DataFrame([[1, 2], [3, 4]], index=index, columns=columns)
    >>> where_row(('R1', 'r1'), df, skip_rows=0)
    (2, 2)
    >>> where_row([('R1', 'r1'), ('R1', 'r2')], df, skip_rows=0)
    (2, 3)
    """
    skip_rows += df.columns.nlevels - 1
    name1, name2 = _fetch_names(name)
    assert df.index.nlevels >= len(name1)
    rows = list(
        zip(*[df.index.get_level_values(-i) for i in range(1, len(name1) + 1)][::-1])
    )
    if name1 in rows and name2 in rows:
        idx1 = rows.index(name1) + 1
        idx2 = rows.index(name2) + 1
        return (
            skip_rows + idx1,
            skip_rows + idx2,
        )


def where(
    name: Name | NameRange, df: pd.DataFrame, skip_rows: int = 3, skip_cols: int = 0
) -> Optional[SpreadsheetRange2d]:
    """
    Return an index of the name in an index of a df

    >>> import pandas as pd
    >>> index = pd.MultiIndex.from_tuples([('R1', 'r1'), ('R1', 'r2')])
    >>> columns = pd.MultiIndex.from_tuples([('C1', 'c1'), ('C1', 'c2')])
    >>> df = pd.DataFrame([[1, 2], [3, 4]], index=index, columns=columns)

    # Searching for a column
    >>> where(('C1', 'c1'), df, skip_rows=0, skip_cols=0)
    (1, 2, 9999, 2)

    # Searching for a row
    >>> where(('R1', 'r1'), df, skip_rows=0, skip_cols=0)
    (2, 1, 2, 9999)
    """
    skip_rows += df.columns.nlevels - 1
    skip_cols += df.index.nlevels - 1
    name1, name2 = _fetch_names(name)

    if df.columns.nlevels >= len(name1):
        cols = list(
            zip(
                *[df.columns.get_level_values(-i) for i in range(1, len(name1) + 1)][
                    ::-1
                ]
            )
        )
        if name1 in cols and name2 in cols:
            idx1 = cols.index(name1) + 1
            idx2 = cols.index(name2) + 1
            return (skip_rows, skip_cols + idx1, 9999, skip_cols + idx2)

    if df.index.nlevels >= len(name1):
        rows = list(
            zip(
                *[df.index.get_level_values(-i) for i in range(1, len(name1) + 1)][::-1]
            )
        )
        if name1 in rows and name2 in rows:
            idx1 = rows.index(name1) + 1
            idx2 = rows.index(name2) + 1
            return (skip_rows + idx1, skip_cols, skip_rows + idx2, 9999)


def cols(r0, c0, r1, c1):
    return list(range(c0, c1 + 1))


def rows(r0, c0, r1, c1):
    return list(range(r0, r1 + 1))
