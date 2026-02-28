# -*- coding: utf-8 -*-

import logging
from traceback import format_exc
import xlsxwriter
from datetime import datetime, date

import pandas as pd
from pandas import isnull, notnull

from .dom import Cell

_default_formats = {
    # Page header
    "title": {
        "font_size": 24,
        "bold": True,
        "align": "left",
        "text_wrap": False,
    },
    "description": {
        "font_size": 9,
        "align": "left",
        "text_wrap": False,
    },
    # Table header
    "table_index_column_names": {
        "font_size": 11,
        "align": "left",
        "valign": "vcenter",
        "text_wrap": False,
        "bold": True,
    },
    "table_column_names": {
        "font_size": 11,
        "align": "center",
        "valign": "vcenter",
        "text_wrap": True,
        "bold": True,
    },
    # Table body
    "table_index_cells": {
        "font_size": 11,
        "align": "left",
        "valign": "vcenter",
        "text_wrap": False,
        "bold": True,
    },
    "table_cells": {
        "font_size": 9,
        "align": "center",
    },
    # Mixins
    "link_mixin": {"font_color": "blue", "underline": True},
    "date_mixin": {"num_format": "yyyy-mm-dd"},
    "datetime_mixin": {"num_format": "yyyy-mm-dd hh:mm"},
    "float_mixin": {"num_format": "#,##0.00"},
    "text_mixin": {"align": "left"},
}

_default_config = {
    "group_border_style": 1,
}


class MergedArea:
    """
    Represents rectangular area of merged cells on a worksheet
    """

    def __init__(self, first_row, first_col, cell):
        self._first_row = first_row
        self._first_col = first_col
        self._last_row = first_row
        self._last_col = first_col
        self._cell = cell

    @property
    def cell(self):
        return self._cell

    def extend(self, num_rows, num_cols):
        self._last_row += num_rows
        self._last_col += num_cols

    @property
    def coords(self):
        return [
            self._first_row,
            self._first_col,
            self._last_row,
            self._last_col,
        ]

    def cell_is_equal_to(self, cell):
        return str(self._cell) == str(cell)


class Renderer:
    """
    Renderer doesn't calculate anything. It only renders, paints and does other formatting.
    """

    @property
    def pages(self):
        return self._pages

    def __init__(self, filename, formats={}, config={}):
        self._filename = filename
        self._pages = []

        # Take default formats, update them with user settings.
        self._formats = _default_formats
        for format_name, format_styles in formats.items():
            self._formats[format_name].update(format_styles)
        assert {
            "title",
            "description",
            "table_index_column_names",
            "table_column_names",
            "table_index_cells",
            "table_cells",
            "link_mixin",
            "float_mixin",
            "text_mixin",
            "datetime_mixin",
            "date_mixin",
        } <= set(self._formats.keys())

        # Take default config, update it with user settings.
        self._config = _default_config
        for k, v in config:
            self._config[k] = v

    def add_page(self, page):
        self._pages.append(page)

    def render_pages(self, skip_errors=False):
        """
        skip_errors: if True, only report errors but continue execution.
        """
        # Open the workbook.
        book = xlsxwriter.Workbook(self._filename)
        # Initialize formats in this book.
        formats = {
            name: book.add_format(props)
            for name, props in self._formats.items()
        }
        # Put pages on sheets of the book.
        for page in self.pages:
            try:
                self._render_page(page, book, formats)
            except:
                if skip_errors:
                    exc_str = format_exc()
                    logging.error(f"Failed to process page {page}:\n{exc_str}")
                    continue
                raise

        book.close()

    def _render_page(self, page, book, formats):
        # Create a worksheet.
        sheet = book.add_worksheet(page.name)

        # Put header and navigation on the page.
        frame = page.table.frame
        if not isinstance(frame.index, pd.MultiIndex):
            frame.index = pd.MultiIndex.from_arrays(
                [frame.index.values], names=[frame.index.name]
            )
        index_shift = len(frame.index.names)
        header_shift = 4

        sheet.write(0, 0, page.header.title, formats["title"])
        sheet.write(1, 0, page.header.description, formats["description"])
        if page.parent is not None:
            sheet.write_url(
                2, 0, f"internal:'{page.parent.name}'!A1", string="Go back"
            )

        sheet.freeze_panes(header_shift, index_shift)

        # Set column widths.
        if page.table.column_widths is not None:
            for column_num, column_width in enumerate(
                page.table.column_widths
            ):
                sheet.set_column(column_num, column_num, column_width)
        # Hide hidden columns.
        for column_num in page.table.hidden_columns:
            sheet.set_column(column_num, column_num, options={"hidden": True})

        # Write table header:
        # columns part
        n_column_levels = len(frame.columns.names)
        # Initialize merged areas list for columns.
        merged_areas = [None] * len(frame.columns.names)
        cell_format = self._formats["table_column_names"].copy()
        for col_num, column_names in enumerate(frame.columns):
            if n_column_levels == 1:
                column_names = [column_names]
            value_have_changed_on_level = None
            for row_num, column_name in enumerate(column_names):
                # Initialize merged areas on first column.
                if col_num == 0:
                    merged_areas[row_num] = MergedArea(
                        header_shift + row_num - 1,
                        index_shift + col_num,
                        column_name,
                    )
                    continue
                # If cell value is equal to one stored in MergedArea for this column,
                # and no cells have changed on prevous columns,
                # then just extend MergedArea for this column by one row.
                if value_have_changed_on_level is None and merged_areas[
                    row_num
                ].cell_is_equal_to(column_name):
                    merged_areas[row_num].extend(0, 1)
                    continue
                # Otherwise, write the current MergedArea and create new one.
                if value_have_changed_on_level is None:
                    value_have_changed_on_level = row_num
                sheet.merge_range(
                    *(merged_areas[row_num].coords),
                    "",
                    # book.add_cell_format),
                )
                self._write_cell(
                    *(merged_areas[row_num].coords[:2]),
                    merged_areas[row_num].cell,
                    sheet,
                    book,
                    cell_format,
                )
                merged_areas[row_num] = MergedArea(
                    header_shift + row_num - 1,
                    index_shift + col_num,
                    column_name,
                )
        # Flush final merged areas.
        for area in merged_areas:
            sheet.merge_range(*(area.coords), "")
            self._write_cell(
                *(area.coords[:2]), area.cell, sheet, book, cell_format
            )
        header_shift += n_column_levels - 1

        # index part
        for col_num, index_name in enumerate(frame.index.names):
            sheet.write(
                header_shift - 1,
                col_num,
                index_name,
                formats["table_index_column_names"],
            )

        # Initialize merged areas list for indices.
        merged_areas = [None] * len(frame.index.names)

        group_level = (
            page.table.group_level
            if page.table.group_level is not None
            else -1
        )

        # Put data on the page.
        # Write rows:
        for row_num, (row_index, row_values) in enumerate(frame.iterrows()):
            # Write indices:
            value_have_changed_on_level = None
            for col_num, index_cell in enumerate(row_index):
                # Initialize merged areas on first row.
                if row_num == 0:
                    merged_areas[col_num] = MergedArea(
                        header_shift, col_num, index_cell
                    )
                    continue
                # If cell value is equal to one stored in MergedArea for this column,
                # and no cells have changed on prevous columns,
                # then just extend MergedArea for this column by one row.
                if value_have_changed_on_level is None and merged_areas[
                    col_num
                ].cell_is_equal_to(index_cell):
                    merged_areas[col_num].extend(1, 0)
                    continue
                # Otherwise, write the current MergedArea and create new one.
                if value_have_changed_on_level is None:
                    value_have_changed_on_level = col_num
                # if group have closed, apply a special style.
                cell_format = self._formats["table_index_cells"].copy()
                if (
                    value_have_changed_on_level is not None
                    and value_have_changed_on_level <= group_level
                ):
                    cell_format.update(
                        {"bottom": self._config["group_border_style"]}
                    )
                sheet.merge_range(
                    *(merged_areas[col_num].coords),
                    "",
                    book.add_format(cell_format),
                )
                self._write_cell(
                    *(merged_areas[col_num].coords[:2]),
                    merged_areas[col_num].cell,
                    sheet,
                    book,
                    cell_format,
                )
                merged_areas[col_num] = MergedArea(
                    header_shift + row_num, col_num, index_cell
                )

            # Write cells:
            for col_num, cell in enumerate(row_values):
                # if group on level 0 have closed, apply a special style.
                cell_format = self._formats["table_cells"].copy()
                if (
                    value_have_changed_on_level is not None
                    and value_have_changed_on_level <= group_level
                ):
                    cell_format.update(
                        {"top": self._config["group_border_style"]}
                    )
                self._write_cell(
                    header_shift + row_num,
                    index_shift + col_num,
                    cell,
                    sheet,
                    book,
                    cell_format,
                )

        # Flush final merged areas.
        for area in merged_areas:
            sheet.merge_range(*(area.coords), "")
            self._write_cell(
                *(area.coords[:2]),
                area.cell,
                sheet,
                book,
                self._formats["table_index_cells"],
            )

        # Draw bottom line to close the table.
        for col_num in range(len(frame.index.names) + len(frame.columns)):
            sheet.write(
                header_shift + row_num + 1,
                col_num,
                None,
                book.add_format({"top": self._config["group_border_style"]}),
            )

        # Add an autofilter
        if page.table.autofilter:
            sheet.autofilter(
                header_shift - 1, 0, header_shift + row_num + 1, col_num
            )

        # Add charts
        for i, spec in enumerate(page.table.charts):
            chart = book.add_chart(
                {"type": spec._type, "subtype": spec._subtype}
            )
            for series in spec.y:
                chart.add_series({"values": series, "categories": spec.x})
            sheet.insert_chart(
                header_shift + row_num + 3 + 20 * i, len(row_index), chart
            )

        # Add conditional formatting
        for i, spec in enumerate(page.table.conditional_formatting):
            sheet.conditional_format(*spec.range, spec.rules)

    def _write_cell(self, row_num, col_num, cell, sheet, book, default_format):
        # For empty cells only apply the format to the cell.
        if isnull(cell):
            value = None
        elif not isinstance(cell, Cell):
            value = cell
        else:
            value = cell.value

        # Assemble cell format:
        format = default_format.copy()
        # values:
        if isinstance(value, datetime):
            format.update(self._formats["datetime_mixin"])
        elif isinstance(value, date):
            format.update(self._formats["date_mixin"])
        elif isinstance(value, float):
            format.update(self._formats["float_mixin"])
        elif isinstance(value, str):
            if value.startswith("="):
                format.update(self._formats["float_mixin"])
            else:
                format.update(self._formats["text_mixin"])

        # color:
        if hasattr(cell, "color") and notnull(cell.color):
            format.update({"bg_color": cell.color})
        # Write cell to the sheet.
        if hasattr(cell, "link") and cell.link is not None:
            link = f"internal:'{cell.link}'!A1"
            format.update(self._formats["link_mixin"])
            format = book.add_format(format)
            sheet.write_url(
                row_num,
                col_num,
                url=link,
                cell_format=format,
                string=str(value),
            )
        elif isinstance(value, str) and value.startswith("="):
            format = book.add_format(format)
            sheet.write_formula(
                row_num, col_num, cell_format=format, formula=value
            )
        else:
            format = book.add_format(format)
            sheet.write(row_num, col_num, value, format)
