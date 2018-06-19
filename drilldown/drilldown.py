# -*- coding: utf-8 -*-
import xlsxwriter
from pandas import isnull, notnull

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
    def groups(self):
        return self._groups

    def __init__(self, frame, groups):
        self._frame = frame
        self._groups = groups


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

    # def _get_value(self):
    #     raise NotImplemented()

    def _get_string(self):
        raise NotImplemented()

    def _get_color(self):
        raise NotImplemented()


class Renderer:
    """
    Renderer doesn't calculate anything. It only renders, paints and does other formatting.
    """
    def __init__(self, filename, formats, column_widths):
        self._filename = filename
        self._column_widths = column_widths
        assert {'title', 'description',
                'table_index_column_names', 'table_column_names',
                'table_index_cells', 'table_cells'} <= set(formats.keys())
        self._formats = formats

    def render_pages(self, pages):
        # Open the workbook.
        book = xlsxwriter.Workbook(self._filename)
        # Initialize formats in this book.
        formats = {name: book.add_format(props)
                   for name, props in self._formats.items() }
        # Put pages on sheets of the book.
        for page in pages:
            self._render_page(page, book, formats)
        book.close()

    def _render_page(self, page, book, formats):
        # Create a worksheet.
        sheet = book.add_worksheet(page.name)

        # Put header and navigation on the page.
        frame = page.table.frame
        index_shift = len(frame.index.names)
        header_shift = 4

        sheet.write(0, 0, page.header.title, formats['title'])
        sheet.write(1, 0, page.header.description, formats['description'])

        sheet.freeze_panes(header_shift, index_shift)

        # Set column widths.
        for column_num, column_width in enumerate(self._column_widths):
            sheet.set_column(column_num, column_num, column_width)

        # Write table header:
        # index part
        for col_num, index_name in enumerate(frame.index.names):
            sheet.write(header_shift-1,
                        col_num,
                        index_name,
                        formats['table_index_column_names'])
        # columns part
        for col_num, column_name in enumerate(frame.columns):
            sheet.write(header_shift-1,
                        index_shift+col_num,
                        column_name,
                        formats['table_column_names'])

        # Put data on the page.
        # Write rows:
        for row_num, (row_index, row_values) in enumerate(frame.iterrows()):
            # write indices
            for col_num, index_value in enumerate(row_index):
                sheet.write(header_shift+row_num,
                            col_num,
                            index_value,
                            formats['table_index_cells'])
            # write cells
            for col_num, cell in enumerate(row_values):
                if isnull(cell):
                    continue
                # Add cell color to the formatting.
                cell_format = self._formats['table_cells'].copy()
                if hasattr(cell, 'color') and notnull(cell.color):
                    cell_format.update({'bg_color': cell.color})
                # Write cell to the sheet.
                sheet.write(header_shift+row_num,
                            index_shift+col_num,
                            str(cell),
                            book.add_format(cell_format))
