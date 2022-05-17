# coding: utf-8
# Python conversion of Calc.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/
import sys
from enum import IntFlag
import numbers
from typing import Iterable, Tuple, cast, overload
from com.sun.star._pyi.table.x_cell_range import XCellRange
from com.sun.star.awt import Point
from com.sun.star.container import XIndexAccess
from com.sun.star.container import XNamed
from com.sun.star.frame import XComponentLoader
from com.sun.star.frame import XController
from com.sun.star.frame import XFrame
from com.sun.star.frame import XModel
from com.sun.star.lang import XComponent
from com.sun.star.rendering import ViewState
from com.sun.star.sheet.CellInsertMode import RIGHT as IM_RIGHT, DOWN as IM_DOWN
from com.sun.star.sheet.CellDeleteMode import LEFT as DM_LEFT, UP as DM_UP
from com.sun.star.sheet import XCellRangeAddressable
from com.sun.star.sheet import XCellRangeMovement
from com.sun.star.sheet import XViewFreezable
from com.sun.star.sheet import XSpreadsheet
from com.sun.star.sheet import XSpreadsheetDocument
from com.sun.star.sheet import XSpreadsheetView
from com.sun.star.sheet import XViewPane
from com.sun.star.sheet.FillDateMode import FILL_DATE_DAY
from com.sun.star.table import CellAddress
from com.sun.star.table import CellRangeAddress
from com.sun.star.table import XColumnRowRange
from com.sun.star.table import XCell
from com.sun.star.table.CellContentType import (
    EMPTY as CCT_EMPTY,
    VALUE as CCT_VALUE,
    TEXT as CCT_TEXT,
    FORMULA as CCT_FORMULA,
)
from com.sun.star.uno import Exception as UnoException

from ..utils import lo as mLo
from ..utils import info as mInfo
from ..utils import gui as mGui
from ..utils import props as mProps
from ..utils.gen_util import ArgsHelper

NameVal = ArgsHelper.NameValue

if sys.version_info >= (3, 10):
    from typing import Union
else:
    from typing_extensions import Union


class Calc:
    # for headers and footers
    HF_LEFT = 0
    HF_CENTER = 1
    HF_RIGHT = 2

    # for zooming, Use GUI.ZoomEnum

    # for border decoration (bitwise composition is possible)
    class BorderEnum(IntFlag):
        TOP_BORDER = 0x01
        BOTTOM_BORDER = 0x02
        LEFT_BORDER = 0x04
        RIGHT_BORDER = 0x08

    # largest value used in XCellSeries.fillSeries
    MAX_VALUE = 0x7FFFFFFF

    # use a better name when date mode doesn't matter
    NO_DATE = FILL_DATE_DAY

    # some hex values for commonly used colors
    BLACK = 0x000000
    WHITE = 0xFFFFFF

    RED = 0xFF0000
    GREEN = 0x00FF00
    BLUE = 0x0000FF
    YELLOW = 0xFFFF00
    ORANGE = 0xFFA500

    DARK_BLUE = 0x003399
    LIGHT_BLUE = 0x99CCFF
    PALE_BLUE = 0xD6EBFF

    CELL_POS = Point(3, 4)

    # --------------- document methods ------------------

    @classmethod
    def open_doc(
        cls, fnm: str, loader: XComponentLoader
    ) -> XSpreadsheetDocument | None:
        doc = mLo.Lo.open_doc(fnm=fnm, loader=loader)
        if doc is None:
            print("Document is null")
            return None
        return cls.get_ss_doc(doc)

    @staticmethod
    def get_ss_doc(doc: XComponent) -> XSpreadsheetDocument | None:
        if not mInfo.Info.is_doc_type(doc_type=mLo.Lo.CALC_SERVICE, obj=doc):
            print("Not a spreadsheet doc; closing")
            mLo.Lo.close_doc(doc=doc)
            return None

        ss_doc = mLo.Lo.qi(XSpreadsheetDocument, doc)
        if ss_doc is None:
            print("Not a spreadsheet doc; closing")
            mLo.Lo.close_doc(doc=doc)
            return None

        return ss_doc

    @staticmethod
    def create_doc(loader: XComponentLoader) -> XSpreadsheetDocument | None:
        doc = mLo.Lo.create_doc(doc_type="scalc", loader=loader)
        return mLo.Lo.qi(XSpreadsheetDocument, doc)
        # XSpreadsheetDocument does not inherit XComponent!

    # ------------------------ sheet methods -------------------------

    @staticmethod
    def _get_sheet_index(doc: XSpreadsheetDocument, index: int) -> XSpreadsheet | None:
        """return the spreadsheet with the specified index (0-based)"""
        sheets = doc.getSheets()
        sheet = None
        try:
            xsheets_idx = mLo.Lo.qi(XIndexAccess, sheets)
            sheet = mLo.Lo.qi(XSpreadsheet, xsheets_idx.getByIndex(index))
        except Exception:
            print(f"Could not access spreadsheet: {index}")
        return sheet

    @staticmethod
    def _get_sheet_name(
        doc: XSpreadsheetDocument, sheet_name: str
    ) -> XSpreadsheet | None:
        """return the spreadsheet with the specified index (0-based)"""
        sheets = doc.getSheets()
        sheet = None
        try:
            sheet = mLo.Lo.qi(XSpreadsheet, sheets.getByName(sheet_name))
        except Exception:
            print(f"Could not access spreadsheet: '{sheet_name}'")
        return sheet

    @overload
    @staticmethod
    def get_sheet(doc: XSpreadsheetDocument, index: int) -> XSpreadsheet | None:
        ...

    @overload
    @staticmethod
    def get_sheet(doc: XSpreadsheetDocument, sheet_name: str) -> XSpreadsheet | None:
        ...

    @classmethod
    def get_sheet(cls, *args, **kwargs) -> XSpreadsheet | None:
        ordered_keys = ("first", "second")
        kargs = {}
        kargs["first"] = kwargs.get("doc", None)
        if "index" in kwargs:
            kargs["second"] = kwargs["index"]
        elif "sheet_name" in kwargs:
            kargs["second"] = kwargs["sheet_name"]
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        k_len = len(kargs)
        if k_len != 2:
            print("invalid number of arguments for get_sheet()")
            return
        if isinstance(kargs["first"], int):
            return cls._get_sheet_index(kargs["first"], kargs["second"])
        return cls._get_sheet_name(kargs["first"], kargs["second"])

    @staticmethod
    def insert_sheet(
        doc: XSpreadsheetDocument, name: str, idx: int
    ) -> XSpreadsheet | None:
        sheets = doc.getSheets()
        sheet = None
        try:
            sheets.insertNewByName(name, idx)
            sheet = mLo.Lo.qi(XSpreadsheet, sheets.getByName(name))
        except Exception as e:
            print("Could not insert sheet:")
            print(f"    {e}")
        return sheet

    @staticmethod
    def _remove_sheet_name(doc: XSpreadsheetDocument, sheet_name: str) -> bool:
        sheets = doc.getSheets()
        try:
            sheets.removeByName(sheet_name)
            return True
        except Exception:
            print(f"Could not remove sheet: {sheet_name}")
        return False

    @classmethod
    def _remove_sheet_index(cls, doc: XSpreadsheetDocument, index: int) -> bool:
        sheets = doc.getSheets()
        try:
            xsheets_idx = mLo.Lo.qi(XIndexAccess, sheets)
            sheet = mLo.Lo.qi(XSpreadsheet, xsheets_idx.getByIndex(index))
            sheet_name = cls.get_sheet_name(sheet)
            if sheet_name is None:
                return False
            sheets.removeByName(sheet_name)
            return True
        except Exception:
            print(f"Could not remove sheet: {sheet_name}")
        return False

    @overload
    @staticmethod
    def remove_sheet(doc: XSpreadsheetDocument, sheet_name: str) -> bool:
        ...

    @overload
    @staticmethod
    def remove_sheet(doc: XSpreadsheetDocument, index: int) -> bool:
        ...

    @overload
    @classmethod
    def remove_sheet(cls, *args, **kwargs) -> bool:
        ordered_keys = ("first", "second")
        kargs = {}
        kargs["first"] = kwargs.get("doc", None)
        if "index" in kwargs:
            kargs["second"] = kwargs["index"]
        elif "sheet_name" in kwargs:
            kargs["second"] = kwargs["sheet_name"]
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        k_len = len(kargs)
        if k_len != 2:
            print("invalid number of arguments for get_sheet()")
            return
        if isinstance(kargs["first"], int):
            return cls._remove_sheet_index(kargs["first"], kargs["second"])
        return cls._remove_sheet_name(kargs["first"], kargs["second"])

    @staticmethod
    def move_sheet(doc: XSpreadsheetDocument, name: str, idx: int) -> bool:
        sheets = doc.getSheets()
        num_sheets = len(sheets.getElementNames())
        if idx < 0 or idx >= num_sheets:
            print(f"Index {idx} is out of range.")
            return False
        sheets.moveByName(name, idx)
        return True

    @staticmethod
    def get_sheet_names(doc: XSpreadsheetDocument) -> Tuple[str, ...]:
        sheets = doc.getSheets()
        return sheets.getElementNames()

    @staticmethod
    def get_sheet_name(sheet: XSpreadsheet) -> str | None:
        xnamed = mLo.Lo.qi(XNamed, sheet)
        if xnamed is None:
            print("Could not access spreadsheet name")
            return None
        return xnamed.getName()

    @staticmethod
    def set_sheet_name(sheet: XSpreadsheet, name: str) -> None:
        xnamed = mLo.Lo.qi(XNamed, sheet)
        if xnamed is None:
            print("Could not access spreadsheet")
            return
        xnamed.setName(name)

    # ----------------- view methods --------------------------

    @staticmethod
    def get_controller(doc: XSpreadsheetDocument) -> XController | None:
        model = mLo.Lo.qi(XModel, doc)
        if model is None:
            return None
        return model.getCurrentController()

    @classmethod
    def zoom_value(cls, doc: XSpreadsheetDocument, value: int) -> None:
        ctrl = cls.get_controller(doc)
        if ctrl is None:
            return
        mProps.Props.set_property(
            prop_set=ctrl, name="ZoomType", value=mGui.GUI.ZoomEnum.BY_VALUE
        )
        mProps.Props.set_property(prop_set=ctrl, name="ZoomValue", value=value)

    @classmethod
    def zoom(cls, doc: XSpreadsheetDocument, type: mGui.GUI.ZoomEnum) -> None:
        ctrl = cls.get_controller(doc)
        if ctrl is None:
            return
        mProps.Props.set_property(prop_set=ctrl, name="ZoomType", value=type)

    @staticmethod
    def get_view(doc: XSpreadsheetDocument) -> XSpreadsheetView | None:
        return mLo.Lo.qi(XSpreadsheetView, doc)

    @classmethod
    def set_active_sheet(cls, doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> None:
        ss_view = cls.get_view(doc)
        if ss_view is None:
            return
        ss_view.setActiveSheet(sheet)

    @classmethod
    def get_active_sheet(cls, doc: XSpreadsheetDocument) -> XSpreadsheet | None:
        ss_view = cls.get_view(doc)
        if ss_view is None:
            return
        return ss_view.getActiveSheet()

    @classmethod
    def freeze(cls, doc: XSpreadsheetDocument, num_cols: int, num_rows: int) -> None:
        ctrl = cls.get_controller(doc)
        if ctrl is None:
            return
        xfreeze = mLo.Lo.qi(XViewFreezable, ctrl)
        xfreeze.freezeAtPosition(num_cols, num_rows)

    @classmethod
    def freeze_cols(cls, doc: XSpreadsheetDocument, num_cols: int) -> None:
        cls.freeze(doc=doc, num_cols=num_cols, num_rows=0)

    @classmethod
    def freeze_rows(cls, doc: XSpreadsheetDocument, num_rows: int) -> None:
        cls.freeze(doc=doc, num_cols=0, num_rows=num_rows)

    @overload
    @staticmethod
    def goto_cell(cell_name: str, doc: XSpreadsheetDocument) -> None:
        ...

    @overload
    @staticmethod
    def goto_cell(cell_name: str, frame: XFrame) -> None:
        ...

    @classmethod
    def goto_cell(cls, *args, **kwargs) -> None:
        ordered_keys = ("first", "second")
        kargs = {}
        kargs["first"] = kwargs.get("cell_name", None)
        if "doc" in kwargs:
            kargs["second"] = kwargs["doc"]
        elif "frame" in kwargs:
            kargs["second"] = kwargs["frame"]
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        k_len = len(kargs)
        if k_len != 2:
            print("invalid number of arguments for goto_cell()")
            return
        doc = mLo.Lo.qi(XSpreadsheetDocument, kargs["second"])
        if doc is not None:
            frame = cls.get_controller(doc).getFrame()
        else:
            frame = kargs["second"]
        props = mProps.Props.make_props(ToPoint=kargs["first"])
        mLo.Lo.dispatch_cmd(cmd="GoToCell", props=props, frame=frame)

    @classmethod
    def split_window(cls, doc: XSpreadsheetDocument, cell_name: str) -> None:
        frame = cls.get_controller(doc).getFrame()
        cls.goto_cell(cell_name=cell_name, frame=frame)
        props = mProps.Props.make_props(ToPoint=cell_name)
        mLo.Lo.dispatch_cmd(cmd="SplitWindow", props=props, frame=frame)

    @overload
    @staticmethod
    def get_selected_addr(doc: XSpreadsheetDocument) -> CellRangeAddress | None:
        ...

    @overload
    @staticmethod
    def get_selected_addr(model: XModel) -> CellRangeAddress | None:
        ...

    @classmethod
    def get_selected_addr_model(cls, *args, **kwargs) -> CellRangeAddress | None:
        ordered_keys = ("first",)
        kargs = {}
        if "doc" in kwargs:
            kargs["first"] = kwargs["doc"]
        elif "model" in kwargs:
            kargs["first"] = kwargs["model"]
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        k_len = len(kargs)
        if k_len != 1:
            print("invalid number of arguments for get_selected_addr_model()")
            return

        doc = mLo.Lo.qi(XSpreadsheetDocument, kargs["first"])
        if doc is not None:
            model = mLo.Lo.qi(XModel, doc)
        else:
            model = cast(XModel, kargs["first"])
        if model is None:
            print("No document model found")
            return None
        ra = mLo.Lo.qi(XCellRangeAddressable, model.getCurrentSelection())
        if ra is None:
            print("No range address found")
            return None
        return ra.getRangeAddress()

    @classmethod
    def get_selected_cell_addr(cls, doc: XSpreadsheetDocument) -> CellRangeAddress:
        cr_addr = cls.get_selected_addr(doc=doc)
        addr = None
        if cls.is_single_cell_range(cr_addr):
            sheet = cls.get_active_sheet(doc)
            cell = cls.get_cell(
                sheet=sheet, column=cr_addr.StartColumn, row=cr_addr.StartRow
            )
            addr = cls.get_cell_address(cell)
        return addr

    # -------------------- view data methods ---------------------------------

    @staticmethod
    def get_view_panes(doc: XSpreadsheetDocument) -> Iterable[XViewPane] | None:
        con = mLo.Lo.qi(XIndexAccess, doc)
        if con is None:
            print("Could not access the view pane container")
            return None
        if con.getCount() == 0:
            print("No view panes found")
            return None

        panes = []
        for i in range(con.getCount()):
            try:
                panes.append(mLo.Lo.qi(XViewPane, con.getByIndex(i)))
            except UnoException:
                print(f"Could not get view pane {i}")
        return panes

    @classmethod
    def get_view_data(cls, doc: XSpreadsheetDocument) -> str:
        ctrl = cls.get_controller(doc)
        return str(ctrl.getViewData())

    @classmethod
    def set_view_data(cls, doc: XSpreadsheetDocument, view_data: str) -> None:
        ctrl = cls.get_controller(doc)
        ctrl.restoreViewData(view_data)

    @classmethod
    def get_view_states(cls, doc: XSpreadsheetDocument) -> Iterable[ViewState] | None:
        """
        Extract the view states for all the sheets from the view data.
        The states are returned as an array of ViewState objects.

        The view data string has the format:
            100/60/0;0;tw:879;0/4998/0/1/0/218/2/0/0/4988/4998

        The view state info starts after the third ";", the fourth entry.
        The view state for each sheet is separated by ";"s

        Based on a post by user Hanya to:
        https://forum.openoffice.org/en/forum/viewtopic.php?f=45&t=29195&p=133202&hilit=getViewData#p133202
        """
        ctrl = cls.get_controller(doc)
        if ctrl is None:
            return None

        view_data = str(ctrl.getViewData())
        view_parts = view_data.split(";")
        p_len = len(view_parts)
        if p_len < 4:
            print("No sheet view states found in view data")
            return None
        states = []
        for i in range(3, p_len):
            states.append(ViewState(view_parts[i]))
        return states

    @classmethod
    def set_view_states(
        cls, doc: XSpreadsheetDocument, states: Iterable[ViewState]
    ) -> None:
        """
        Update the sheet state part of the view data, which starts as
        the 4th entry in the view data string
        """
        ctrl = cls.get_controller(doc)
        if ctrl is None:
            return
        view_data = str(ctrl.getViewData())
        view_parts = view_data.split(";")
        p_len = len(view_parts)
        if p_len < 4:
            print("No sheet view states found in view data")
            return None

        vd_new = []
        for i in range(4):
            vd_new.append(view_parts[i])

        for state in states:
            vd_new.append(str(state))
        s_data = ";".join(vd_new)
        print(s_data)
        ctrl.restoreViewData(s_data)

    # ----------- insert/remove rows, columns, cells ---------------

    @staticmethod
    def insert_row(sheet: XSpreadsheet, idx: int) -> None:
        cr_range = mLo.Lo.qi(XColumnRowRange, sheet)
        rows = cr_range.getRows()
        rows.insertByIndex(idx, 1)  # add 1 row at idx position

    @staticmethod
    def delete_row(sheet: XSpreadsheet, idx: int) -> None:
        cr_range = mLo.Lo.qi(XColumnRowRange, sheet)
        rows = cr_range.getRows()
        rows.removeByIndex(idx, 1)  # remove 1 row at idx position

    @staticmethod
    def insert_column(sheet: XSpreadsheet, idx: int) -> None:
        cr_range = mLo.Lo.qi(XColumnRowRange, sheet)
        cols = cr_range.getColumns()
        cols.insertByIndex(idx, 1)  # add 1 column at idx position

    @staticmethod
    def delete_column(sheet: XSpreadsheet, idx: int) -> None:
        cr_range = mLo.Lo.qi(XColumnRowRange, sheet)
        cols = cr_range.getColumns()
        cols.removeByIndex(idx, 1)  # remove 1 row at idx position

    @classmethod
    def insert_cells(
        cls, sheet: XSpreadsheet, cell_range: XCellRange, is_shift_right: bool
    ) -> None:
        mover = mLo.Lo.qi(XCellRangeMovement, sheet)
        addr = cls.get_address(cell_range)
        if is_shift_right:
            mover.insertCells(addr, IM_RIGHT)
        else:
            mover.insertCells(addr, IM_DOWN)

    @classmethod
    def delete_cells(
        cls, sheet: XSpreadsheet, cell_range: XCellRange, is_shift_right: bool
    ) -> None:
        mover = mLo.Lo.qi(XCellRangeMovement, sheet)
        addr = cls.get_address(cell_range)
        if is_shift_right:
            mover.removeRange(addr, DM_LEFT)
        else:
            mover.removeRange(addr, DM_UP)

    # ----------- set/get values in cells ------------------
    @staticmethod
    def _set_val_by_cell(value: object, cell: XCell) -> None:
        if isinstance(value, numbers.Number):
            cell.setValue(float(value))
        elif isinstance(value, str):
            cell.setFormula(str(value))
        else:
            print(f"Value is not a number or string: {value}")

    @classmethod
    def _set_val_by_cell_name(
        cls, value: object, sheet: XSpreadsheet, cell_name: str
    ) -> None:
        pos = cls.get_cell_position(cell_name)
        cls._set_val_by_col_row(value=value, sheet=sheet, column=pos.X, row=pos.Y)

    @classmethod
    def _set_val_by_col_row(
        cls, value: object, sheet: XSpreadsheet, column: int, row: int
    ) -> None:
        cell = cls.get_cell(sheet=sheet, column=column, row=row)
        cls._set_val_by_cell(value=value, cell=cell)

    @overload
    @staticmethod
    def set_val(value: object, cell: XCell) -> None:
        ...

    @overload
    @staticmethod
    def set_val(value: object, sheet: XSpreadsheet, cell_name: str) -> None:
        ...

    @overload
    @staticmethod
    def set_val(value: object, sheet: XSpreadsheet, column: int, row: int) -> None:
        ...

    @classmethod
    def set_val(cls, *args, **kwargs) -> None:
        ordered_keys = ("first", "second", "third", "fourth")
        kargs = {}
        kargs["first"] = kwargs.get("value", None)
        if "cell" in kwargs:
            kargs["second"] = kwargs["cell"]
        elif "sheet" in kwargs:
            kargs["second"] = kwargs["sheet"]
        elif "cell_name" in kwargs:
            kargs["third"] = kwargs["cell_name"]
        elif "column" in kwargs:
            kargs["third"] = kwargs["column"]
        elif "row" in kwargs:
            kargs["fourth"] = kwargs["row"]
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        k_len = len(kargs)
        if k_len < 2 or k_len > 4:
            print("invalid number of arguments for set_val()")
            return
        if k_len == 2:
            cls._set_val_by_cell(value=kargs["first"], cell=kargs["second"])
        elif k_len == 3:
            cls._set_val_by_cell_name(
                value=kargs["first"], sheet=kargs["second"], cell_name=kargs["third"]
            )
        elif k_len == 4:
            cls._set_val_by_col_row(
                value=kargs["first"],
                sheet=kargs["second"],
                column=kargs["third"],
                row=kargs["fourth"],
            )

    @staticmethod
    def convert_to_double(val: object) -> float:
        if val is None:
            print("Value is null; using 0")
            return 0
        try:
            return float(val)
        except ValueError:
            print(f"Could not convert {val} to double; using 0")
            return 0

    @staticmethod
    def get_type_string(cell: XCell) -> str:
        t = cell.getType()
        if t == CCT_EMPTY:
            return "EMPTY"
        if t == CCT_VALUE:
            return "VALUE"
        if t == CCT_TEXT:
            return "TEXT"
        if t == CCT_FORMULA:
            return "FORMULA"
        print("Unknown cell type")
        return "??"

    @classmethod
    def _get_val_by_cell(cls, cell: XCell, column: int, row: int) -> object | None:
        t = cell.getType()
        if t == CCT_EMPTY:
            return None
        if t == CCT_VALUE:
            return float(cell.getValue())
        if t == CCT_TEXT or t == CCT_FORMULA:
            return cell.getFormula()
        print("Unknown cell type; returning None")
        return None

    @classmethod
    def _get_val_by_col_row(
        cls, sheet: XSpreadsheet, column: int, row: int
    ) -> object | None:
        xcell = cls.get_cell(sheet=sheet, column=column, row=row)
        return cls._get_val_by_cell(cell=xcell, column=column, row=row)

    @classmethod
    def _get_val_by_cell_name(
        cls, sheet: XSpreadsheet, cell_name: str
    ) -> object | None:
        pos = cls.get_cell_position(cell_name)
        return cls._get_val_by_col_row(sheet=sheet, column=pos.X, row=pos.Y)
    
    @classmethod
    def _get_val_by_cell_addr(
        cls, sheet: XSpreadsheet, addr: CellAddress
    ) -> object | None:
        if addr is None:
            return None
        return cls._get_val_by_col_row(sheet=sheet, column=addr.Column, row=addr.Row)

    @overload
    @staticmethod
    def get_val(sheet: XSpreadsheet, addr: CellAddress) -> object | None:
        ...

    @overload
    @staticmethod
    def get_val(sheet: XSpreadsheet, cell_name: str) -> object | None:
        ...

    @overload
    @staticmethod
    def get_val(sheet: XSpreadsheet, column: int, row: int) -> object | None:
        ...

    @overload
    @staticmethod
    def get_val(cell: XCell, column: int, row: int) -> object | None:
        ...

    @classmethod
    def get_val(cls, *args, **kwargs) -> object | None:
        ordered_keys = ("first", "second", "third")
        def get_kwargs() -> dict:
            kargs = {}
            key = "sheet"
            if key in kwargs:
                kargs["first"] = kwargs[key]

            key = "cell"
            if key in kwargs:
                kargs["first"] = kwargs[key]

            if kargs['first'] is None:
                return kargs
            
            key = "addr"
            if key in kwargs:
                kargs["second"] = kwargs[key]
            
            key = "cell_name"
            if key in kwargs:
                kargs["second"] = kwargs[key]
            
            key = "column"
            if key in kwargs:
                kargs["second"] = kwargs[key]

            if kargs['second'] is None:
                return kargs
            key = "row"
            if key in kwargs:
                kargs["third"] = kwargs[key]
            return kargs
            
        kargs = get_kwargs()
        
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        
        k_len = len(kargs)
        if k_len < 2 or k_len > 4:
            print("invalid number of arguments for set_val()")
            return
        
        first_arg =  mLo.Lo.qi(XSpreadsheet, kargs["first"])
        if first_arg is None:
            # can only be: get_val(cell: XCell, column: int, row: int)
            return cls._get_val_by_cell(cell=kargs['first'], column=kargs['second'], row=kargs['third'])

        if k_len == 2:
            # get_val(sheet: XSpreadsheet, addr: CellAddress) or
            # get_val(sheet: XSpreadsheet, cell_name: str)
            if isinstance(kargs['second'], str):
                # get_val(sheet: XSpreadsheet, cell_name: str)
                return cls._get_val_by_cell_name(sheet=kargs['first'], cell_name=kargs['second'])
            return cls._get_val_by_cell_addr(sheet=kargs['first'], addr=kargs['second'])

        if k_len == 3:
            # get_val(sheet: XSpreadsheet, column: int, row: int) 
            return cls._get_val_by_col_row(sheet=kargs['first'], column=kargs['second'], row=kargs['third'])
        return None

    @overload
    @staticmethod
    def get_num(sheet: XSpreadsheet, cell_name: str) -> float:...
    
    @overload
    @staticmethod
    def get_num(sheet: XSpreadsheet, addr: CellAddress) -> float:...
    
    @overload
    @staticmethod
    def get_num(sheet: XSpreadsheet, column: int, row: int) -> float:...
    
    @classmethod
    def get_num(cls, *args, **kwargs) -> float:
        ordered_keys = ("first", "second", "third")
        def get_kwargs() -> dict:
            kargs = {}
            kargs['first']= kwargs.get("sheet", None)
            if kargs['first'] is None:
                return kargs
            
            key = "cell_name"
            if key in kwargs:
                kargs["second"] = kwargs[key]
            
            key = "addr"
            if key in kwargs:
                kargs["second"] = kwargs[key]
            
            key = "column"
            if key in kwargs:
                kargs["second"] = kwargs[key]

            if kargs['second'] is None:
                return kargs
            key = "row"
            if key in kwargs:
                kargs["third"] = kwargs[key]
            return kargs
            
        kargs = get_kwargs()
        
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        
        k_len = len(kargs)
        if k_len < 2 or k_len > 3:
            print("invalid number of arguments for get_num()")
            return 0
        if k_len == 3:
            return cls.convert_to_double(cls.get_val(sheet=kargs['first'], column=kargs['second'], row=kargs['third']))
        if k_len == 2:
            if isinstance(kargs["second"], str):
                return cls.convert_to_double(cls.get_val(sheet=kargs['first'], cell_name=kargs['second']))
            return cls.convert_to_double(cls.get_val(sheet=kargs['first'], addr=kargs['second']))
        return 0
    
    @overload
    @staticmethod
    def get_string(sheet: XSpreadsheet, cell_name: str) -> str | None:...
    
    @overload
    @staticmethod
    def get_string(sheet: XSpreadsheet, cell_name: str) -> str | None:...
    
    @overload
    @staticmethod
    def get_string(sheet: XSpreadsheet, addr: CellAddress) -> str | None:...
    
    @overload
    @staticmethod
    def get_string(sheet: XSpreadsheet, column: int, row: int) -> str | None:...
    
    @classmethod
    def get_string(cls, *args, **kwargs) -> str | None:
        ordered_keys = ("first", "second", "third")
        def get_kwargs() -> dict:
            kargs = {}
            kargs['first']= kwargs.get("sheet", None)
            if kargs['first'] is None:
                return kargs
            
            key = "cell_name"
            if key in kwargs:
                kargs["second"] = kwargs[key]
            
            key = "addr"
            if key in kwargs:
                kargs["second"] = kwargs[key]
            
            key = "column"
            if key in kwargs:
                kargs["second"] = kwargs[key]

            if kargs['second'] is None:
                return kargs
            key = "row"
            if key in kwargs:
                kargs["third"] = kwargs[key]
            return kargs
            
        kargs = get_kwargs()
        
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        
        k_len = len(kargs)
        if k_len < 2 or k_len > 3:
            print("invalid number of arguments for get_string()")
            return None
        if k_len == 3:
            return str(cls.get_val(sheet=kargs['first'], column=kargs['second'], row=kargs['third']))
        if k_len == 2:
            if isinstance(kargs["second"], str):
                return str(cls.get_val(sheet=kargs['first'], cell_name=kargs['second']))
            return str(cls.get_val(sheet=kargs['first'], addr=kargs['second']))
        return None
    
    # ----------- set/get values in 2D array ------------------