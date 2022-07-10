# coding: utf-8
"""
Calc Named Events.
"""
from __future__ import annotations
from typing import NamedTuple


class CalcNamedEvent(NamedTuple):
    """
    Named events for :py:class:`~.office.calc.Calc` class
    """

    DOC_OPENING = "calc_doc_opening"
    """Doc Opening Spreadsheet see :py:meth:`Calc.open_doc() <.office.calc.Calc.open_doc>`"""
    DOC_OPENED = "calc_doc_opened"
    """Doc Opened Spreadsheet see :py:meth:`Calc.open_doc() <.office.calc.Calc.open_doc>`"""

    DOC_SS = "calc_doc_ss"
    """Doc Get Spreadsheet see :py:meth:`Calc.get_ss_doc() <.office.calc.Calc.get_ss_doc>`"""

    DOC_CREATING = "calc_doc_creating"
    """Doc Creating see :py:meth:`Calc.create_doc() <.office.calc.Calc.create_doc>`"""
    DOC_CREATED = "calc_doc_created"
    """Doc Created see :py:meth:`Calc.create_doc() <.office.calc.Calc.create_doc>`"""

    DOC_SAVING = "calc_doc_saving"
    """Doc Saving Spreadsheet document see :py:meth:`Calc.save_doc() <.office.calc.Calc.save_doc>`"""
    DOC_SAVED = "calc_doc_saved"
    """Doc Saving Spreadsheet document see :py:meth:`Calc.save_doc() <.office.calc.Calc.save_doc>`"""

    SHEET_GETTING = "calc_sheet_getting"
    """Sheet Getting see :py:meth:`Calc.get_sheet() <.office.calc.Calc.get_sheet>`"""
    SHEET_GET = "calc_sheet_get"
    """Sheet Get see :py:meth:`Calc.get_sheet() <.office.calc.Calc.get_sheet>`"""

    SHEET_INSERTING = "calc_sheet_inserting"
    """Sheet Inserting see :py:meth:`Calc.insert_sheet() <.office.calc.Calc.insert_sheet>`"""
    SHEET_INSERTED = "calc_sheet_inserted:"
    """Sheet Inserted see :py:meth:`Calc.insert_sheet() <.office.calc.Calc.insert_sheet>`"""

    SHEET_REMOVING = "calc_sheet_removing"
    """Sheet Removing see :py:meth:`Calc.remove_sheet() <.office.calc.Calc.remove_sheet>`"""
    SHEET_REMOVED = "calc_sheet_removed"
    """Sheet Removed see :py:meth:`Calc.remove_sheet() <.office.calc.Calc.remove_sheet>`"""

    SHEET_MOVING = "calc_sheet_moving"
    """Sheet Moving see :py:meth:`Calc.move_sheet() <.office.calc.Calc.move_sheet>`"""
    SHEET_MOVED = "calc_sheet_moved"
    """Sheet Moved see :py:meth:`Calc.move_sheet() <.office.calc.Calc.move_sheet>`"""

    SHEET_ACTIVATING = "calc_sheet_activating"
    """Sheet Activating see :py:meth:`Calc.set_active_sheet() <.office.calc.Calc.set_active_sheet>`"""
    SHEET_ACTIVATED = "calc_sheet_activated"
    """Sheet Activated see :py:meth:`Calc.set_active_sheet() <.office.calc.Calc.set_active_sheet>`"""

    SHEET_ROW_INSERTING = "calc_sheet_row_inserting"
    """Sheet row inserting see :py:meth:`Calc.insert_row() <.office.calc.Calc.insert_row>`"""
    SHEET_ROW_INSERTED = "calc_sheet_row_inserted"
    """Sheet row inserted see :py:meth:`Calc.insert_row() <.office.calc.Calc.insert_row>`"""

    SHEET_ROW_DELETING = "calc_sheet_row_deleting"
    """Sheet row deleting see :py:meth:`Calc.delete_row() <.office.calc.Calc.delete_row>`"""
    SHEET_ROW_DELETED = "calc_sheet_row_deleted"
    """Sheet row deleted see :py:meth:`Calc.delete_row() <.office.calc.Calc.delete_row>`"""

    SHEET_COL_INSERTING = "calc_sheet_col_inserting"
    """Sheet column inserting see :py:meth:`Calc.insert_column() <.office.calc.Calc.insert_column>`"""
    SHEET_COL_INSERTED = "calc_sheet_col_inserted"
    """Sheet column inserted see :py:meth:`Calc.insert_column() <.office.calc.Calc.insert_column>`"""

    SHEET_COL_DELETING = "calc_sheet_col_deleting"
    """Sheet column deleting see :py:meth:`Calc.delete_column() <.office.calc.Calc.delete_column>`"""
    SHEET_COL_DELETED = "calc_sheet_col_deleted"
    """Sheet column deleted see :py:meth:`Calc.delete_column() <.office.calc.Calc.delete_column>`"""

    SHEET_COL_WIDTH_SETTING = "calc_sheet_col_width_setting"
    """Sheet setting column width see :py:meth:`Calc.set_col_width() <.office.calc.Calc.set_col_width>`"""
    SHEET_COL_WIDTH_SET = "calc_sheet_col_width_set"
    """Sheet set column width see :py:meth:`Calc.set_col_width() <.office.calc.Calc.set_col_width>`"""

    SHEET_ROW_HEIGHT_SETTING = "calc_sheet_row_height_setting"
    """Sheet setting row height see :py:meth:`Calc.set_row_height() <.office.calc.Calc.set_row_height>`"""
    SHEET_ROW_HEIGHT_SET = "calc_sheet_row_height_set"
    """Sheet set row height see :py:meth:`Calc.set_row_height() <.office.calc.Calc.set_row_height>`"""

    CELLS_INSERTING = "calc_cells_inserting"
    """Cells Inserting see :py:meth:`Calc.insert_cells() <.office.calc.Calc.insert_cells>`"""
    CELLS_INSERTED = "calc_cells_inserted"
    """Cells Inserted see :py:meth:`Calc.insert_cells() <.office.calc.Calc.insert_cells>`"""

    CELLS_DELETING = "calc_cells_deleting"
    """Cells Deleting see :py:meth:`Calc.delete_cells() <.office.calc.Calc.delete_cells>`"""
    CELLS_DELETED = "calc_cells_deleted"
    """Cells Deleted see :py:meth:`Calc.delete_cells() <.office.calc.Calc.delete_cells>`"""

    CELLS_CLEARING = "calc_cells_clearing"
    """Cells Clearing see :py:meth:`Calc.clear_cells() <.office.calc.Calc.clear_cells>`"""
    CELLS_CLEARED = "calc_cells_cleared"
    """Cells Cleared see :py:meth:`Calc.clear_cells() <.office.calc.Calc.clear_cells>`"""

    CELLS_BORDER_REMOVING = "calc_cells_border_removing"
    """Cells Border Removing see :py:meth:`Calc.remove_border() <.office.calc.Calc.remove_border>`"""
    CELLS_BORDER_REMOVED = "calc_cells_border_removed"
    """Cells Border Removed see :py:meth:`Calc.remove_border() <.office.calc.Calc.remove_border>`"""

    CELLS_BORDER_ADDING = "calc_cells_border_adding"
    """Cells Border Adding see :py:meth:`Calc.add_border() <.office.calc.Calc.add_border>`"""
    CELLS_BORDER_ADDED = "calc_cells_border_added"
    """Cells Border Added see :py:meth:`Calc.add_border() <.office.calc.Calc.add_border>`"""

    CELLS_HIGH_LIGHTING = "calc_cells_high_lighting"
    """Cells Highlighting see :py:meth:`Calc.highlight_range() <.office.calc.Calc.highlight_range>`"""
    CELLS_HIGH_LIGHTED = "calc_cells_high_lighted"
    """Cells Highlighted see :py:meth:`Calc.highlight_range() <.office.calc.Calc.highlight_range>`"""
