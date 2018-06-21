# -*- coding: utf-8 -*-

"""
These classes represent structure of a book.

Presentation details should be avoided to allow writing
to different output formats (i.e. a spreadsheet or a website).
"""


class Header:
    # Title is short summary of what is on the page.
    @property
    def title(self):
        return self._title

    # Description explains details: how values are computed, etc.
    @property
    def description(self):
        return self._description

    def __init__(self, title, description):
        self._title = title
        self._description = description


class Navbar:
    # Top-level (horizontal) structure of the book.
    @property
    def links(self):
        return self._links

    # Top-down (vertical) levels representing drillable structure of the book.
    @property
    def breadcrumbs(self):
        return self._breadcrumbs

    def __init__(self, links, breadcrumbs):
        self._links = links
        self._breadcrumbs = breadcrumbs


class Table:
    # Dataframe with `columns`, `index`, `values`.
    @property
    def frame(self):
        return self._frame

    # Groups are folded using outlines (with total row).
    # Formatting visually separates groups between each other.
    @property
    def group_level(self):
        return self._group_level

    # Visual properties.
    @property
    def column_widths(self):
        return self._column_widths

    @property
    def hidden_columns(self):
        return self._hidden_columns

    def __init__(self, frame, group_level=None, column_widths=None,
                 hidden_columns=[]):
        self._frame = frame
        self._group_level = group_level
        self._column_widths = column_widths
        self._hidden_columns = hidden_columns


class Page:
    # Page name -- used for references.
    @property
    def name(self):
        return self._name

    # Page parent (if needed).
    @property
    def parent(self):
        return self._parent

    # Page structural elements: header, navbar, data table.
    @property
    def header(self):
        return self._header

    @property
    def navbar(self):
        return self._navbar

    @property
    def table(self):
        return self._table

    def __repr__(self):
        return f"Page(name={self.name})"

    def __init__(self, name, header, table, navbar=None, parent=None):
        self._name = name
        self._parent = parent
        self._header = header
        self._navbar = navbar
        self._table = table

    def set_navbar(self, navbar):
        """ Navbar can be set afterwards"""
        self._navbar = navbar


class Cell:
    """
    Cell has few abstract properties describing its content: string, background color.

    They do not depend on the exact backend. Renderer should take care
    of converting them to backend parlance.
    """

    # @property
    # def value(self):
    #     return self._get_value()

    @property
    def string(self):
        return self._get_string()

    def __str__(self):
        return self._get_string()

    @property
    def color(self):
        return self._get_color()

    @property
    def link(self):
        return self._get_link()

    # def _get_value(self):
    #     raise NotImplemented()

    def _get_string(self):
        # User must implement this method to use cell.
        raise NotImplemented()

    def _get_color(self):
        # By default cell has no color. User may override this.
        return None

    def _get_link(self):
        # By default cell has no link. User may override this.
        return None

    # These two are implemented to allow `Cell`s in pandas multiindices.
    # Pandas converts all index values to `Categorical`s with `ordered=True`,
    # so we provide means to compare cells.
    def __le__(self, other):
        return str(self) <= str(other)

    def __lt__(self, other):
        return str(self) < str(other)

    def __eq__(self, other):
        return str(self) == str(other)

    def __hash__(self):
        return str(self).__hash__()
