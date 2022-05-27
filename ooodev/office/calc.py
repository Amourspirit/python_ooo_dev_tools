# coding: utf-8
# Python conversion of Calc.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/
# region Imports
from __future__ import annotations
from enum import IntFlag, Enum
import numbers
import re
from typing import Any, List, Tuple, cast, overload, Sequence, TYPE_CHECKING
import uno

from com.sun.star.awt import Point
from com.sun.star.container import XIndexAccess
from com.sun.star.container import XNamed
from com.sun.star.frame import XModel
from com.sun.star.lang import XComponent
from com.sun.star.lang import Locale
from com.sun.star.sheet.GeneralFunction import (
    NONE as GF_NONE,
    AUTO as GF_AUTO,
    SUM as GF_SUM,
    COUNT as GF_COUNT,
    AVERAGE as GF_AVERAGE,
    MAX as GF_MAX,
    MIN as GF_MIN,
    PRODUCT as GF_PRODUCT,
    COUNTNUMS as GF_COUNTNUMS,
    STDEV as GF_STDEV,
    STDEVP as GF_STDEVP,
    VAR as GF_VAR,
    VARP as GF_VARP,
)
from com.sun.star.sheet import SolverConstraint  # struct
from com.sun.star.sheet.SolverConstraintOperator import (
    LESS_EQUAL as SCO_LESS_EQUAL,
    EQUAL as SCO_EQUAL,
    GREATER_EQUAL as SCO_GREATER_EQUAL,
    INTEGER as SCO_INTEGER,
    BINARY as SCO_BINARY,
)
from com.sun.star.sheet import XCellAddressable
from com.sun.star.sheet import XCellRangeData
from com.sun.star.sheet import XCellRangeAddressable
from com.sun.star.sheet import XCellRangeMovement
from com.sun.star.sheet import XCellSeries
from com.sun.star.sheet import XDataPilotTable
from com.sun.star.sheet import XDataPilotTablesSupplier
from com.sun.star.sheet import XFunctionAccess
from com.sun.star.sheet import XFunctionDescriptions
from com.sun.star.sheet import XHeaderFooterContent
from com.sun.star.sheet import XRecentFunctions
from com.sun.star.sheet import XScenario
from com.sun.star.sheet import XScenariosSupplier
from com.sun.star.sheet import XSpreadsheet
from com.sun.star.sheet import XSpreadsheetDocument
from com.sun.star.sheet import XSpreadsheetView
from com.sun.star.sheet import XSheetAnnotationAnchor
from com.sun.star.sheet import XSheetAnnotationsSupplier
from com.sun.star.sheet import XSheetCellRange
from com.sun.star.sheet import XSheetOperation
from com.sun.star.sheet import XUsedAreaCursor
from com.sun.star.sheet import XViewPane
from com.sun.star.sheet import XViewFreezable
from com.sun.star.sheet.CellDeleteMode import LEFT as DM_LEFT, UP as DM_UP
from com.sun.star.sheet.CellInsertMode import RIGHT as IM_RIGHT, DOWN as IM_DOWN
from com.sun.star.sheet.FillDateMode import FILL_DATE_DAY
from com.sun.star.style import XStyle
from com.sun.star.table import BorderLine2  # struct
from com.sun.star.table import TableBorder2  # struct
from com.sun.star.table import XColumnRowRange
from com.sun.star.table import XCellRange
from com.sun.star.table.CellContentType import (
    EMPTY as CCT_EMPTY,
    VALUE as CCT_VALUE,
    TEXT as CCT_TEXT,
    FORMULA as CCT_FORMULA,
)
from com.sun.star.text import XSimpleText
from com.sun.star.uno import Exception as UnoException
from com.sun.star.util import NumberFormat  # const
from com.sun.star.util import XNumberFormatsSupplier
from com.sun.star.util import XNumberFormatTypes

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue
    from com.sun.star.beans import XPropertySet
    from com.sun.star.frame import XComponentLoader
    from com.sun.star.frame import XController
    from com.sun.star.frame import XFrame

    # from com.sun.star.sheet import CellAnnotation
    from com.sun.star.sheet import FunctionArgument  # struct
    from com.sun.star.sheet import XSheetAnnotation
    from com.sun.star.sheet import XDataPilotTables
    from com.sun.star.sheet import XGoalSeek
    from com.sun.star.sheet import XSheetCellCursor
    from com.sun.star.sheet import XSolver
    from com.sun.star.table import CellAddress
    from com.sun.star.table import CellRangeAddress
    from com.sun.star.table import XCell
    from com.sun.star.text import XText
    from com.sun.star.util import XSearchable
    from com.sun.star.util import XSearchDescriptor


from ..utils import lo as mLo
from ..utils import info as mInfo
from ..utils import gui as mGui
from ..utils import props as mProps
from ..utils.gen_util import ArgsHelper, TableHelper, Util as GenUtil
from ..utils import enum_helper
from ..utils.color import CommonColor
from ..utils import view_state as mViewState

NameVal = ArgsHelper.NameValue
# endregion Imports


class Calc:
    # region classes
    # for headers and footers
    class HeaderFooter:
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

    class CellTypeEnum(str, Enum):
        EMPTY = "EMPTY"
        VALUE = "VALUE"
        TEXT = "TEXT"
        FORMULA = "FORMULA"
        UNKNOWN = "UNKNOWN"

    class GeneralFunction:
        __typename__ = "com.sun.star.sheet.GeneralFunction"
        NONE = GF_NONE
        AUTO = GF_AUTO
        SUM = GF_SUM
        COUNT = GF_COUNT
        AVERAGE = GF_AVERAGE
        MAX = GF_MAX
        MIN = GF_MIN
        PRODUCT = GF_PRODUCT
        COUNTNUMS = GF_COUNTNUMS
        STDEV = GF_STDEV
        STDEVP = GF_STDEVP
        VAR = GF_VAR
        VARP = GF_VARP

    setattr(GeneralFunction, "__new__", enum_helper.uno_enum_class_new)

    class SolverConstraintOperator:
        __typename__ = "com.sun.star.sheet.SolverConstraintOperator"
        LESS_EQUAL = SCO_LESS_EQUAL
        EQUAL = SCO_EQUAL
        GREATER_EQUAL = SCO_GREATER_EQUAL
        INTEGER = SCO_INTEGER
        BINARY = SCO_BINARY

    setattr(SolverConstraintOperator, "__new__", enum_helper.uno_enum_class_new)

    # endregion classes

    # region Constants
    # largest value used in XCellSeries.fillSeries
    MAX_VALUE = 0x7FFFFFFF

    # use a better name when date mode doesn't matter
    NO_DATE = FILL_DATE_DAY

    CELL_POS = Point(3, 4)

    _rx_cell = re.compile(r"([a-zA-Z]+)([0-9]+)")

    # endregion Constants

    # region --------------- document methods --------------------------

    @classmethod
    def open_doc(cls, fnm: str, loader: XComponentLoader) -> XSpreadsheetDocument | None:
        doc = mLo.Lo.open_doc(fnm=str(fnm), loader=loader)
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

    # endregion ------------ document methods ------------------

    # region --------------- sheet methods -----------------------------

    # region    get_sheet()
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
    def _get_sheet_name(doc: XSpreadsheetDocument, sheet_name: str) -> XSpreadsheet | None:
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
        """
        Gets a sheet of spreadsheet document

        Args:
            doc (XSpreadsheetDocument): Spreadsheet document
            index (int): Zero based index of spreadsheet

        Returns:
            XSpreadsheet | None: Spreadsheet at index if found; Otherwise, False
        """
        ...

    @overload
    @staticmethod
    def get_sheet(doc: XSpreadsheetDocument, sheet_name: str) -> XSpreadsheet | None:
        """
        Gets a sheet of spreadsheet document

        Args:
            doc (XSpreadsheetDocument): Spreadsheet document
            sheet_name (str): Name of spreadsheet

        Returns:
            XSpreadsheet | None: Spreadsheet with matching name if found; Otherwise, False
        """
        ...

    @classmethod
    def get_sheet(cls, *args, **kwargs) -> XSpreadsheet | None:
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('doc','index', 'sheet_name')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_sheet() got an unexpected keyword argument")
            ka[1] = kwargs.get("doc", None)
            keys = ("index", "sheet_name")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            return ka

        if count != 2:
            raise TypeError("get_sheet() got an invalid numer of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if isinstance(kargs[2], int):
            return cls._get_sheet_index(kargs[1], kargs[2])
        return cls._get_sheet_name(kargs[1], kargs[2])

    # endregion get_sheet()

    @staticmethod
    def insert_sheet(doc: XSpreadsheetDocument, name: str, idx: int) -> XSpreadsheet | None:
        """
        Inserts a spreadsheet into document.

        Args:
            doc (XSpreadsheetDocument): Spreadsheet document
            name (str): Name of sheet to insert
            idx (int): zero-based index position of the sheet to insert

        Returns:
            XSpreadsheet | None: The newly inserted sheet on success; Othwrwise, None
        """
        sheets = doc.getSheets()
        sheet = None
        try:
            sheets.insertNewByName(name, idx)
            sheet = mLo.Lo.qi(XSpreadsheet, sheets.getByName(name))
        except Exception as e:
            print("Could not insert sheet:")
            print(f"    {e}")
        return sheet

    # region    remove_sheet()

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
            print(f"Could not remove sheet at index: {index}")
        return False

    @overload
    @staticmethod
    def remove_sheet(doc: XSpreadsheetDocument, sheet_name: str) -> bool:
        """
        Removes a sheet from document

        Args:
            doc (XSpreadsheetDocument): Spreadsheet document
            sheet_name (str): Name of sheet to remove

        Returns:
            bool: True of sheet was removed; Otherwise, False
        """
        ...

    @overload
    @staticmethod
    def remove_sheet(doc: XSpreadsheetDocument, index: int) -> bool:
        """
        Removes a sheet from document

        Args:
            doc (XSpreadsheetDocument): Spreadsheet document
            index (int): Zero based index of sheet to remove.

        Returns:
            bool: True of sheet was removed; Otherwise, False
        """
        ...

    @classmethod
    def remove_sheet(cls, *args, **kwargs) -> bool:
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('doc', 'index', 'sheet_name')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("remove_sheet() got an unexpected keyword argument")
            ka[1] = kwargs.get("doc", None)
            keys = ("index", "sheet_name")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            return ka

        if count != 2:
            raise TypeError("remove_sheet() got an invalid numer of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if isinstance(kargs[2], int):
            return cls._remove_sheet_index(kargs[1], kargs[2])
        return cls._remove_sheet_name(kargs[1], kargs[2])

    # endregion remove_sheet()

    @staticmethod
    def move_sheet(doc: XSpreadsheetDocument, name: str, idx: int) -> bool:
        """
        Moves a sheet in a spreadsheet document

        Args:
            doc (XSpreadsheetDocument): Spreadsheet document.
            name (str): Name of sheet to move
            idx (int): The zero based index to move sheet into.

        Returns:
            bool: True on success; Otherwise, False
        """
        sheets = doc.getSheets()
        num_sheets = len(sheets.getElementNames())
        if idx < 0 or idx >= num_sheets:
            print(f"Index {idx} is out of range.")
            return False
        sheets.moveByName(name, idx)
        return True

    @staticmethod
    def get_sheet_names(doc: XSpreadsheetDocument) -> Tuple[str, ...]:
        """
        Gets names of all existing spreadsheets in the spreadsheet document.

        Args:
            doc (XSpreadsheetDocument): Document to get sheets of

        Returns:
            Tuple[str, ...]: Tuple of sheet names.
        """
        sheets = doc.getSheets()
        return sheets.getElementNames()

    @staticmethod
    def get_sheet_name(sheet: XSpreadsheet) -> str | None:
        """
        Gets the name of a sheet

        Args:
            sheet (XSpreadsheet): Spreadsheet

        Returns:
            str | None: Name of sheet on success; Otherwise, None
        """
        xnamed = mLo.Lo.qi(XNamed, sheet)
        if xnamed is None:
            print("Could not access spreadsheet name")
            return None
        return xnamed.getName()

    @staticmethod
    def set_sheet_name(sheet: XSpreadsheet, name: str) -> None:
        """
        Sets the name of a spreadsheet.

        Args:
            sheet (XSpreadsheet): Spreadsheet to set name of
            name (str): New name for spreadsheet.
        """
        xnamed = mLo.Lo.qi(XNamed, sheet)
        if xnamed is None:
            print("Could not access spreadsheet")
            return
        xnamed.setName(name)

    # endregion --------------------- sheet methods -------------------------

    # region --------------- view methods ------------------------------

    @staticmethod
    def get_controller(doc: XSpreadsheetDocument) -> XController | None:
        """
        Provides access to the controller which currently controls this model

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document

        Returns:
            XController | None: Controller for Spreadsheet Document
        """
        model = mLo.Lo.qi(XModel, doc)
        if model is None:
            return None
        return model.getCurrentController()

    @classmethod
    def zoom_value(cls, doc: XSpreadsheetDocument, value: int) -> None:
        """
        Sets the zoom level of the Spreadsheet Document

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document
            value (int): Value to set zoom. e.g. 160 set zoom to 160%
        """
        ctrl = cls.get_controller(doc)
        if ctrl is None:
            return
        mProps.Props.set_property(prop_set=ctrl, name="ZoomType", value=mGui.GUI.ZoomEnum.BY_VALUE)
        mProps.Props.set_property(prop_set=ctrl, name="ZoomValue", value=value)

    @classmethod
    def zoom(cls, doc: XSpreadsheetDocument, type: mGui.GUI.ZoomEnum) -> None:
        """
        Zooms spreadsheet document to a specific view.

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document
            type (mGui.GUI.ZoomEnum): Type of Zoom to set.
        """
        ctrl = cls.get_controller(doc)
        if ctrl is None:
            return
        mProps.Props.set_property(prop_set=ctrl, name="ZoomType", value=type)

    @classmethod
    def get_view(cls, doc: XSpreadsheetDocument) -> XSpreadsheetView | None:
        """
        Is the main interface of a SpreadsheetView.

        It manages the active sheet within this view.

        The com.sun.star.sheet.SpreadsheetView service is the spreadsheet's extension
        of the com.sun.star.frame.Controller service and represents a table editing view
        for a spreadsheet document.

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document

        Returns:
            XSpreadsheetView | None: XSpreadsheetView on success; Otherwise, None
        """
        return mLo.Lo.qi(XSpreadsheetView, cls.get_controller(doc))

    @classmethod
    def set_active_sheet(cls, doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> None:
        """
        Sets the active sheet

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document
            sheet (XSpreadsheet): Sheet to set active
        """
        ss_view = cls.get_view(doc)
        if ss_view is None:
            return
        ss_view.setActiveSheet(sheet)

    @classmethod
    def get_active_sheet(cls, doc: XSpreadsheetDocument) -> XSpreadsheet | None:
        """
        Gets the active sheet

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document

        Returns:
            XSpreadsheet | None: Active Sheet if found; Otherwise, None
        """
        ss_view = cls.get_view(doc)
        if ss_view is None:
            return
        return ss_view.getActiveSheet()

    @classmethod
    def freeze(cls, doc: XSpreadsheetDocument, num_cols: int, num_rows: int) -> None:
        """
        Freezes spreadsheet columns and rows

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document
            num_cols (int): Number of columns to freeze
            num_rows (int): Number of rows to freeze
        """
        ctrl = cls.get_controller(doc)
        if ctrl is None:
            return
        if num_cols < 0 or num_rows < 0:
            return
        xfreeze = mLo.Lo.qi(XViewFreezable, ctrl)
        xfreeze.freezeAtPosition(num_cols, num_rows)

    @classmethod
    def freeze_cols(cls, doc: XSpreadsheetDocument, num_cols: int) -> None:
        """
        Freezes spreadsheet columns

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document
            num_cols (int): Number of columns to freeze
        """
        cls.freeze(doc=doc, num_cols=num_cols, num_rows=0)

    @classmethod
    def freeze_rows(cls, doc: XSpreadsheetDocument, num_rows: int) -> None:
        """
        Freezes spreadsheet rows

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document
            num_rows (int): Number of rows to freeze
        """
        cls.freeze(doc=doc, num_cols=0, num_rows=num_rows)

    # region    goto_cell()
    @overload
    @staticmethod
    def goto_cell(cell_name: str, doc: XSpreadsheetDocument) -> None:
        """
        Go to a cell

        Args:
            cell_name (str): Cell Name such as 'B4'
            doc (XSpreadsheetDocument): Spreadsheet Document
        """
        ...

    @overload
    @staticmethod
    def goto_cell(cell_name: str, frame: XFrame) -> None:
        """
        Go to a cell

        Args:
            cell_name (str): Cell Name such as 'B4'
            frame (XFrame): Spreadsheet frame.
        """
        ...

    @classmethod
    def goto_cell(cls, *args, **kwargs) -> None:
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs():
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('cell_name', 'doc', 'frame')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("goto_cell() got an unexpected keyword argument")
            ka[1] = kwargs.get("cell_name", None)
            keys = ("doc", "frame")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            return ka

        if count != 2:
            raise TypeError("set_val() got an invalid numer of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        doc = mLo.Lo.qi(XSpreadsheetDocument, kargs[2])
        if doc is not None:
            frame = cls.get_controller(doc).getFrame()
        else:
            frame = kargs[2]
        props = mProps.Props.make_props(ToPoint=kargs[1])
        mLo.Lo.dispatch_cmd(cmd="GoToCell", props=props, frame=frame)

    # endregion    goto_cell()

    @classmethod
    def split_window(cls, doc: XSpreadsheetDocument, cell_name: str) -> None:
        """
        Splits window

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document
            cell_name (str): Cell to preform split on. e.g. 'C4'
        """
        frame = cls.get_controller(doc).getFrame()
        cls.goto_cell(cell_name=cell_name, frame=frame)
        props = mProps.Props.make_props(ToPoint=cell_name)
        mLo.Lo.dispatch_cmd(cmd="SplitWindow", props=props, frame=frame)

    # region    get_selected_addr()

    @overload
    @staticmethod
    def get_selected_addr(doc: XSpreadsheetDocument) -> CellRangeAddress | None:
        """
        Gets select cell range addresses

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document

        Returns:
            CellRangeAddress | None: Cell range adresses on success; Othwrwise, None
        """
        ...

    @overload
    @staticmethod
    def get_selected_addr(model: XModel) -> CellRangeAddress | None:
        """
        Gets select cell range addresses

        Args:
            model (XModel): model used to access sheet

        Returns:
            CellRangeAddress | None: Cell range adresses on success; Othwrwise, None
        """
        ...

    @classmethod
    def get_selected_addr(cls, *args, **kwargs) -> CellRangeAddress | None:
        ordered_keys = (1,)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('doc', 'model')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_selected_addr() got an unexpected keyword argument")
            keys = ("doc", "model")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            return ka

        if count != 1:
            raise TypeError("get_selected_addr() got an invalid numer of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        doc = mLo.Lo.qi(XSpreadsheetDocument, kargs[1])
        if doc:
            model = mLo.Lo.qi(XModel, doc)
        else:
            # def get_selected_addr(model: XModel)
            model = cast(XModel, kargs[1])

        if model is None:
            print("No document model found")
            return None
        ra = mLo.Lo.qi(XCellRangeAddressable, model.getCurrentSelection())
        if ra is None:
            print("No range address found")
            return None
        return ra.getRangeAddress()

    # endregion  get_selected_addr()

    @classmethod
    def get_selected_cell_addr(cls, doc: XSpreadsheetDocument) -> CellAddress | None:
        """
        Gets the cell address of current selected cell of the active sheet.

        Args:
            doc (XSpreadsheetDocument): Spreadsheet document

        Returns:
            CellAddress | None: CellAddress on success; Otherwise, None

        Note:
            If more then a single cell is selected then ``None`` is returned.

            CellAddress returns Zero-base values.
            For instance: Cell ``B4`` has Column value of ``1`` and Row value of ``3``
        """
        cr_addr = cls.get_selected_addr(doc=doc)
        addr = None
        if cls.is_single_cell_range(cr_addr):
            sheet = cls.get_active_sheet(doc)
            cell = cls.get_cell(sheet=sheet, column=cr_addr.StartColumn, row=cr_addr.StartRow)
            addr = cls.get_cell_address(cell)
        return addr

    # endregion -------------- view methods ----------------------------

    # region --------------- view data methods -------------------------

    @classmethod
    def get_view_panes(cls, doc: XSpreadsheetDocument) -> List[XViewPane] | None:
        """
        represents a pane in a view of a spreadsheet document.

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document

        Returns:
            List[XViewPane] | None: List of XViewPane on success; Otherwise, None

        Notes:
            The com.sun.star.sheet.XViewPane interface's getFirstVisibleColumn(), getFirstVisibleRow(),
            setFirstVisibleColumn() and setFirstVisibleRow() methods query and set the start of
            the exposed area. The getVisibleRange() method returns a com.sun.star.table.
            CellRangeAddress struct describing which cells are shown in the pane.
            Columns or rows that are only partly visible at the right or lower edge of the view
            are not included.
        """
        con = mLo.Lo.qi(XIndexAccess, cls.get_controller(doc))
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
        """
        Gets a set of data that can be used to restore the current view status at
        later time by using ``set_view_data()``

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document

        Returns:
            str: View Data
        """
        ctrl = cls.get_controller(doc)
        return str(ctrl.getViewData())

    @classmethod
    def set_view_data(cls, doc: XSpreadsheetDocument, view_data: str) -> None:
        """
        Restores the view status using the data gotten from a previous call to
        ``get_view_data()``

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document
            view_data (str): Data to restore.
        """
        ctrl = cls.get_controller(doc)
        ctrl.restoreViewData(view_data)

    @classmethod
    def get_view_states(cls, doc: XSpreadsheetDocument) -> List[mViewState.ViewState] | None:
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
            states.append(mViewState.ViewState(view_parts[i]))
        return states

    @classmethod
    def set_view_states(cls, doc: XSpreadsheetDocument, states: Sequence[mViewState.ViewState]) -> None:
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

    # endregion ----------------- view data methods ---------------------------------

    # region --------------- insert/remove rows, columns, cells --------

    @staticmethod
    def insert_row(sheet: XSpreadsheet, idx: int) -> None:
        """
        Inserts a row in spreadsheet

        Args:
            sheet (XSpreadsheet): Spreadsheet
            idx (int): Zero base index of row to insert.
        """
        cr_range = mLo.Lo.qi(XColumnRowRange, sheet)
        rows = cr_range.getRows()
        rows.insertByIndex(idx, 1)  # add 1 row at idx position

    @staticmethod
    def delete_row(sheet: XSpreadsheet, idx: int) -> None:
        """
        Deletes a row from spreadsheet

        Args:
            sheet (XSpreadsheet): Spreadsheet
            idx (int): Zero based index of row to delete
        """
        cr_range = mLo.Lo.qi(XColumnRowRange, sheet)
        rows = cr_range.getRows()
        rows.removeByIndex(idx, 1)  # remove 1 row at idx position

    @staticmethod
    def insert_column(sheet: XSpreadsheet, idx: int) -> None:
        """
        Inserts a column in a spreadsheet.

        Args:
            sheet (XSpreadsheet): Spreadsheet
            idx (int): Zero base index of column to insert.
        """
        cr_range = mLo.Lo.qi(XColumnRowRange, sheet)
        cols = cr_range.getColumns()
        cols.insertByIndex(idx, 1)  # add 1 column at idx position

    @staticmethod
    def delete_column(sheet: XSpreadsheet, idx: int) -> None:
        """
        Delete a column from a spreadsheet

        Args:
            sheet (XSpreadsheet): Spreadsheet
            idx (int): Zero base of index of column to delete
        """
        cr_range = mLo.Lo.qi(XColumnRowRange, sheet)
        cols = cr_range.getColumns()
        cols.removeByIndex(idx, 1)  # remove 1 row at idx position

    @classmethod
    def insert_cells(cls, sheet: XSpreadsheet, cell_range: XCellRange, is_shift_right: bool) -> None:
        """
        Inserts Cells into a spreadsheet

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cell_range (XCellRange): Cell range to insert
            is_shift_right (bool): If True then cell are inserted to the right; Otherwise, inserted down.
        """
        mover = mLo.Lo.qi(XCellRangeMovement, sheet)
        addr = cls.get_address(cell_range)
        if is_shift_right:
            mover.insertCells(addr, IM_RIGHT)
        else:
            mover.insertCells(addr, IM_DOWN)

    @classmethod
    def delete_cells(cls, sheet: XSpreadsheet, cell_range: XCellRange, is_shift_left: bool) -> None:
        """
        Deletes cell in a spreadsheet

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cell_range (XCellRange): Cell range to delete
            is_shift_left (bool): If True then cell are shifted left; Otherwise, cells are shifted up.
        """
        mover = mLo.Lo.qi(XCellRangeMovement, sheet)
        addr = cls.get_address(cell_range)
        if is_shift_left:
            mover.removeRange(addr, DM_LEFT)
        else:
            mover.removeRange(addr, DM_UP)

    # endregion ------------ insert/remove rows, columns, cells -----

    # region --------------- set/get values in cells -------------------
    # region    set_val()
    @staticmethod
    def _set_val_by_cell(value: object, cell: XCell) -> None:
        if isinstance(value, numbers.Number):
            cell.setValue(float(value))
        elif isinstance(value, str):
            cell.setFormula(str(value))
        else:
            print(f"Value is not a number or string: {value}")

    @classmethod
    def _set_val_by_cell_name(cls, value: object, sheet: XSpreadsheet, cell_name: str) -> None:
        pos = cls.get_cell_position(cell_name)
        cls._set_val_by_col_row(value=value, sheet=sheet, column=pos.X, row=pos.Y)

    @classmethod
    def _set_val_by_col_row(cls, value: object, sheet: XSpreadsheet, column: int, row: int) -> None:
        cell = cls.get_cell(sheet=sheet, column=column, row=row)
        cls._set_val_by_cell(value=value, cell=cell)

    @overload
    @staticmethod
    def set_val(value: object, cell: XCell) -> None:
        """
        Sets the value of a cell

        Args:
            value (object): Value for cell
            cell (XCell): Cell to assign value
        """
        ...

    @overload
    @staticmethod
    def set_val(value: object, sheet: XSpreadsheet, cell_name: str) -> None:
        """
        Sets the value of a cell

        Args:
            value (object): Value for cell
            sheet (XSpreadsheet): Spreadsheet
            cell_name (str): Name of cel to set value of such as 'B4'
        """
        ...

    @overload
    @staticmethod
    def set_val(value: object, sheet: XSpreadsheet, column: int, row: int) -> None:
        """
        Sets the value of a cell

        Args:
            value (object): Value for cell
            sheet (XSpreadsheet): Spreadsheet
            column (int): Cell column as zero-based integer
            row (int): Cell row as zero-based integer
        """
        ...

    @classmethod
    def set_val(cls, *args, **kwargs) -> None:
        ordered_keys = (1, 2, 3, 4)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('value', 'cell', 'sheet', 'cell_name', 'column', 'row')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("set_val() got an unexpected keyword argument")
            ka[1] = kwargs.get("value", None)
            keys = ("cell", "sheet")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka

            keys = ("cell_name", "column")
            for key in keys:
                if key in kwargs:
                    ka[3] = kwargs[key]
                    break
            if count == 3:
                return ka
            ka[4] = kwargs.get("row", None)
            return ka

        if not count in (2, 3, 4):
            raise TypeError("set_val() got an invalid numer of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 2:
            cls._set_val_by_cell(value=kargs[1], cell=kargs[2])
        elif count == 3:
            cls._set_val_by_cell_name(value=kargs[1], sheet=kargs[2], cell_name=kargs[3])
        elif count == 4:
            cls._set_val_by_col_row(value=kargs[1], sheet=kargs[2], column=kargs[3], row=kargs[4])

    # endregion    set_val()

    @staticmethod
    def convert_to_float(val: object) -> float:
        """
        Converts value to float

        Args:
            val (object): Value to convert

        Returns:
            float: value converted to float. 0.0 is returned if conversion fails.
        """
        if val is None:
            print("Value is null; using 0")
            return 0.0
        try:
            return float(val)
        except ValueError:
            print(f"Could not convert {val} to double; using 0")
            return 0.0

    convert_to_double = convert_to_float

    @classmethod
    def get_type_string(cls, cell: XCell) -> CellTypeEnum:
        """
        Gets String representing the Type

        Args:
            cell (XCell): _description_

        Returns:
            str: _description_
        """
        t = cell.getType()
        if t == CCT_EMPTY:
            return cls.CellTypeEnum.EMPTY
        if t == CCT_VALUE:
            return cls.CellTypeEnum.VALUE
        if t == CCT_TEXT:
            return cls.CellTypeEnum.TEXT
        if t == CCT_FORMULA:
            return cls.CellTypeEnum.FORMULA
        print("Unknown cell type")
        return cls.CellTypeEnum.UNKNOWN

    # region    get_val()

    @classmethod
    def _get_val_by_cell(cls, cell: XCell) -> object | None:
        t = cell.getType()
        if t == CCT_EMPTY:
            return None
        if t == CCT_VALUE:
            return cls.convert_to_float(cell.getValue())
        if t == CCT_TEXT or t == CCT_FORMULA:
            return cell.getFormula()
        print("Unknown cell type; returning None")
        return None

    @classmethod
    def _get_val_by_col_row(cls, sheet: XSpreadsheet, column: int, row: int) -> object | None:
        xcell = cls.get_cell(sheet=sheet, column=column, row=row)
        return cls._get_val_by_cell(cell=xcell)

    @classmethod
    def _get_val_by_cell_name(cls, sheet: XSpreadsheet, cell_name: str) -> object | None:
        pos = cls.get_cell_position(cell_name)
        return cls._get_val_by_col_row(sheet=sheet, column=pos.X, row=pos.Y)

    @classmethod
    def _get_val_by_cell_addr(cls, sheet: XSpreadsheet, addr: CellAddress) -> object | None:
        if addr is None:
            return None
        return cls._get_val_by_col_row(sheet=sheet, column=addr.Column, row=addr.Row)

    @overload
    @staticmethod
    def get_val(cell: XCell) -> object | None:
        """
        Gets cell value

        Args:
            cell (XCell): cell to get value of

        Returns:
            object | None: Cell value cell has a value; Otherwise, None
        """
        ...

    @overload
    @staticmethod
    def get_val(sheet: XSpreadsheet, addr: CellAddress) -> object | None:
        """
        Get cell value

        Args:
            sheet (XSpreadsheet): Spreadsheet
            addr (CellAddress): Address of cell

        Returns:
            object | None: Cell value cell has a value; Otherwise, None
        """
        ...

    @overload
    @staticmethod
    def get_val(sheet: XSpreadsheet, cell_name: str) -> object | None:
        """
        Gets cell value

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cell_name (str): Name of cell such as 'B4'

        Returns:
            object | None: Cell value cell has a value; Otherwise, None
        """
        ...

    @overload
    @staticmethod
    def get_val(sheet: XSpreadsheet, column: int, row: int) -> object | None:
        """
        Get cell value

        Args:
            sheet (XSpreadsheet): Spreadsheet
            column (int): Cell zero-based column
            row (int): Cell zero-base row

        Returns:
            object | None: Cell value cell has a value; Otherwise, None
        """
        ...

    @classmethod
    def get_val(cls, *args, **kwargs) -> object | None:
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('sheet', 'cell', 'cell_name', 'addr', 'column', 'row')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_val() got an unexpected keyword argument")
            keys = ("sheet", "cell")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            keys = ("addr", "cell_name", "column")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("row", None)
            return ka

        if not count in (1, 2, 3):
            raise TypeError("get_val() got an invalid numer of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        first_arg = mLo.Lo.qi(XSpreadsheet, kargs[1])
        if first_arg is None:
            # can only be: get_val(cell: XCell)
            if count != 1:
                return None
            return cls._get_val_by_cell(cell=kargs[1])

        if count == 2:
            if isinstance(kargs[2], str):
                #   get_val(sheet: XSpreadsheet, cell_name: str)
                return cls._get_val_by_cell_name(sheet=kargs[1], cell_name=kargs[2])

            #   get_val(sheet: XSpreadsheet, addr: CellAddress)
            return cls._get_val_by_cell_addr(sheet=kargs[1], addr=kargs[2])

        if count == 3:
            #   get_val(sheet: XSpreadsheet, column: int, row: int)
            return cls._get_val_by_col_row(sheet=kargs[1], column=kargs[2], row=kargs[3])
        return None

    # endregion get_val()

    # region    get_num()

    # cell: XCell
    @overload
    @staticmethod
    def get_num(cell: XCell) -> float:
        """
        Get cell value a float

        Args:
            cell (XCell): Cell to get value of

        Returns:
            float: Cell value as float. If cell value cannot be converted then 0.0 is returned.
        """
        ...

    @overload
    @staticmethod
    def get_num(sheet: XSpreadsheet, cell_name: str) -> float:
        """
        Gets cell value as float

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cell_name (str): Cell name such as 'B4'

        Returns:
            float: Cell value as float. If cell value cannot be converted then 0.0 is returned.
        """
        ...

    @overload
    @staticmethod
    def get_num(sheet: XSpreadsheet, addr: CellAddress) -> float:
        """
        Gets cell value as float

        Args:
            sheet (XSpreadsheet): Spreadsheet
            addr (CellAddress): Cell Address

        Returns:
            float: Cell value as float. If cell value cannot be converted then 0.0 is returned.
        """
        ...

    @overload
    @staticmethod
    def get_num(sheet: XSpreadsheet, column: int, row: int) -> float:
        """
        Gets cell value as float

        Args:
            sheet (XSpreadsheet): Spreadsheet
            column (int): Cell zero-base column number
            row (int): Cell zero-base row number.

        Returns:
            float: Cell value as float. If cell value cannot be converted then 0.0 is returned.
        """
        ...

    @classmethod
    def get_num(cls, *args, **kwargs) -> float:
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('sheet', 'cell', 'cell_name', 'addr', 'column', 'row')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_num() got an unexpected keyword argument")
            keys = ("sheet", "cell")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            keys = ("cell_name", "addr", "column")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("row", None)
            return ka

        if not count in (1, 2, 3):
            raise TypeError("get_num() got an invalid numer of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            return cls.convert_to_float(cls.get_val(cell=kargs[1]))

        if count == 3:
            return cls.convert_to_float(cls.get_val(sheet=kargs[1], column=kargs[2], row=kargs[3]))
        if count == 2:
            if isinstance(kargs[2], str):
                return cls.convert_to_float(cls.get_val(sheet=kargs[1], cell_name=kargs[2]))
            return cls.convert_to_float(cls.get_val(sheet=kargs[1], addr=kargs[2]))
        return 0.0

    # endregion get_num()

    # region    get_string()
    @overload
    @staticmethod
    def get_string(cell: XCell) -> str:
        """
        Gets the value of a cell as a string.

        Args:
            cell (XCell): Cell to get value of

        Returns:
            str: Cell value as string.
        """
        ...

    @overload
    @staticmethod
    def get_string(sheet: XSpreadsheet, cell_name: str) -> str:
        """
        Gets the value of a cell as a string

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cell_name (str): Name of cell to get the value of such as 'B4'

        Returns:
            str: Cell value as string
        """
        ...

    @overload
    @staticmethod
    def get_string(sheet: XSpreadsheet, addr: CellAddress) -> str:
        """
        Gets the value of a cell as a string

        Args:
            sheet (XSpreadsheet): Spreadsheet
            addr (CellAddress): Cell address

        Returns:
            str: Cell value as string
        """
        ...

    @overload
    @staticmethod
    def get_string(sheet: XSpreadsheet, column: int, row: int) -> str:
        """
        Gets the value of a cell as a string

        Args:
            sheet (XSpreadsheet): Spreadsheet
            column (int): Cell zero-based column number
            row (int): Cell zero-based row number

        Returns:
            str: Cell value as string
        """
        ...

    @classmethod
    def get_string(cls, *args, **kwargs) -> str:
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('cell', 'sheet', 'cell_name', 'addr', 'column', 'row')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_string() got an unexpected keyword argument")
            keys = ("cell", "sheet")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            keys = ("cell_name", "addr", "column")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("row", None)
            return ka

        if not count in (1, 2, 3):
            raise TypeError("get_string() got an invalid numer of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        def convert(obj) -> str:
            if obj is None:
                return ""
            return str(obj)

        if count == 1:
            return convert(cls.get_val(cell=kargs[1]))

        if count == 3:
            return convert(cls.get_val(sheet=kargs[1], column=kargs[2], row=kargs[3]))
        if count == 2:
            if isinstance(kargs[2], str):
                return convert(cls.get_val(sheet=kargs[1], cell_name=kargs[2]))
            return convert(cls.get_val(sheet=kargs[1], addr=kargs[2]))
        return None

    # endregion get_string()

    # endregion ------------ set/get values in cells -----------------

    # region --------------- set/get values in 2D array ----------------

    # region    set_array()
    @classmethod
    def _set_array_doc_addr(
        cls, values: Sequence[Sequence[object]], doc: XSpreadsheetDocument, addr: CellAddress
    ) -> None:
        v_len = len(values)
        if v_len == 0:
            print("Values has not data")
            return
        sheet = cls._get_sheet_index(doc=doc, index=addr.Sheet)
        col_end = addr.Column + (len(values[0]) - 1)
        row_end = addr.Row + (v_len - 1)
        cell_range = cls._get_cell_range_col_row(
            sheet=sheet, start_col=addr.Column, start_row=addr.Row, end_col=col_end, end_row=row_end
        )
        cls.set_cell_range_array(cell_range=cell_range, values=values)

    @overload
    @staticmethod
    def set_array(values: Sequence[Sequence[object]], cell_range: XCellRange) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Sequence[Sequence[object]]): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            cell_range (XCellRange): Range in spreadsheet to insert data
        """
        ...

    @overload
    @staticmethod
    def set_array(values: Sequence[Sequence[object]], sheet: XSpreadsheet, name: str) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Sequence[Sequence[object]]): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            sheet (XSpreadsheet): Spreadsheet
            name (str): Range name such as 'A1:D4' or cell name such as 'B4'

        Notes:
            If ``name`` is a single cell such as ``A1`` then then values are inserter at the named cell
            and expand to the size of the value array.
        """
        ...

    @overload
    @staticmethod
    def set_array(values: Sequence[Sequence[object]], doc: XSpreadsheetDocument, addr: CellAddress) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Sequence[Sequence[object]]): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            doc (XSpreadsheetDocument): Spreadsheet Document
            addr (CellAddress): Address to insert data.
        """
        ...

    @overload
    @staticmethod
    def set_array(
        values: Sequence[Sequence[object]],
        sheet: XSpreadsheet,
        col_start: int,
        row_start: int,
        col_end: int,
        row_end: int,
    ) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Sequence[Sequence[object]]): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            sheet (XSpreadsheet): Spreadsheet
            col_start (int): Zero-base Start Colum
            row_start (int): Zero-base Start Row
            col_end (int): Zero-base End Colum
            row_end (int): Zero-base End Row
        """
        ...

    @classmethod
    def set_array(cls, *args, **kwargs) -> None:
        ordered_keys = (1, 2, 3, 4, 5, 6)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('values', 'cell_range', 'sheet','doc', 'name', 'col_start', 'addr', 'row_start', 'col_end', 'row_end')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("set_array() got an unexpected keyword argument")
            ka[1] = kwargs.get("values", None)
            keys = ("cell_range", "sheet", "doc")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            keys = ("name", "col_start", "addr")
            for key in keys:
                if key in kwargs:
                    ka[3] = kwargs[key]
                    break
            if count == 3:
                return ka
            ka[4] = kwargs.get("row_start", None)
            ka[5] = kwargs.get("col_end", None)
            ka[6] = kwargs.get("row_end", None)
            return ka

        if not count in (2, 3, 6):
            raise TypeError("set_array() got an invalid numer of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 2:
            #  set_array(values: Sequence[Sequence[object]], cell_range: XCellRange)
            cls.set_cell_range_array(cell_range=kargs[2], values=kargs[1])
        if count == 3:
            if isinstance(kargs[3], str):
                # set_array(values: Sequence[Sequence[object]], sheet: XSpreadsheet, name: str)
                if cls.is_cell_range_name(kargs[3]):
                    cls.set_array_range(sheet=kargs[2], range_name=kargs[3], values=kargs[1])
                    return
                else:
                    cls.set_array_cell(sheet=kargs[2], cell_name=kargs[3], values=kargs[1])
                    return
            else:
                cls._set_array_doc_addr(values=kargs[1], doc=kargs[2], addr=kargs[3])
        if count == 6:
            #  def set_array(values: Sequence[Sequence[object]], sheet: XSpreadsheet, col_start: int, row_start: int, col_end:int, row_end: int)
            cell_range = cls._get_cell_range_col_row(
                sheet=kargs[2], start_col=kargs[3], start_row=kargs[4], end_col=kargs[5], end_row=kargs[6]
            )
            cls.set_cell_range_array(cell_range=cell_range, values=kargs[1])
        return

    # endregion set_array()

    @classmethod
    def set_array_range(
        cls,
        sheet: XSpreadsheet,
        range_name: str,
        values: Sequence[Sequence[object]],
    ) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            sheet (XSpreadsheet): _description_
            range_name (str): Range to insert data such as 'A1:E12'
            values (Sequence[Sequence[object]]): A 2-Dimensional array of value such as a list of list or tuple of tuples.
        """
        v_len = len(values)
        if v_len == 0:
            print("Values has not data")
            return
        cell_range = cls.get_cell_range(sheet=sheet, range_name=range_name)
        cls.set_cell_range_array(cell_range=cell_range, values=values)

    @staticmethod
    def set_cell_range_array(cell_range: XCellRange, values: Sequence[Sequence[object]]) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            cell_range (XCellRange): Cell Ranage
            values (Sequence[Sequence[object]]): A 2-Dimensional array of value such as a list of list or tuple of tuples.
        """
        v_len = len(values)
        if v_len == 0:
            print("Values has not data")
            return
        cr_data = mLo.Lo.qi(XCellRangeData, cell_range)
        if cr_data is None:
            return
        cr_data.setDataArray(values)

    @classmethod
    def set_array_cell(cls, sheet: XSpreadsheet, cell_name: str, values: Sequence[Sequence[object]]) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cell_name (str): Cell Name such as 'A1'
            values (Sequence[Sequence[object]]): A 2-Dimensional array of value such as a list of list or tuple of tuples.
        """
        v_len = len(values)
        if v_len == 0:
            print("Values has not data")
            return
        pos = cls.get_cell_position(cell_name)
        col_end = pos.X + (len(values[0]) - 1)
        row_end = pos.Y + (v_len - 1)
        cell_range = cls._get_cell_range_col_row(
            sheet=sheet,
            start_col=pos.X,
            start_row=pos.Y,
            end_col=col_end,
            end_row=row_end,
        )
        cls.set_cell_range_array(cell_range=cell_range, values=values)

    # region get_array()

    @overload
    @staticmethod
    def get_array(cell_range: XCellRange) -> Tuple[Tuple[object, ...], ...]:
        """
        Gets Array of data from a spreadsheet.

        Args:
            cell_range (XCellRange): Cell range that to get data from.

        Returns:
            Tuple[Tuple[object, ...], ...]: Resulting data array.
        """
        ...

    @overload
    @staticmethod
    def get_array(sheet: XSpreadsheet, range_name: str) -> Tuple[Tuple[object, ...], ...]:
        """
        Gets Array of data from a spreadsheet.

        Args:
            sheet (XSpreadsheet): Spreadsheet
            range_name (str): Range of data to get such as "A1:E16"

        Returns:
            Tuple[Tuple[object, ...], ...]: Resulting data array.
        """
        ...

    @classmethod
    def get_array(cls, *args, **kwargs) -> Tuple[Tuple[object, ...], ...]:
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('cell_range', 'sheet','range_name')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_array() got an unexpected keyword argument")
            keys = ("cell_range", "sheet")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            ka[2] = kwargs.get("range_name", None)
            return ka

        if not count in (1, 2):
            raise TypeError("get_array() got an invalid numer of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            cell_range = cast(XCellRange, kargs[1])
        else:
            cell_range = cls.get_cell_range(sheet=kargs[1], range_name=kargs[2])

        cr_data = mLo.Lo.qi(XCellRangeData, cell_range)
        return cr_data.getDataArray()

    # endregion get_array()

    @staticmethod
    def print_array(vals: Sequence[Sequence[object]]) -> None:
        """
        Prints a 2-Dimensional array to terminal

        Args:
            vals (Sequence[Sequence[object]]): A 2-Dimensional array of value such as a list of list or tuple of tuples.
        """
        row_len = len(vals)
        if row_len == 0:
            print("No data in array to print")
            return
        col_len = len(vals[0])
        print(f"Row x Column size: {row_len} x {col_len}")
        for row in vals:
            col_str = "  ".join([str(cell) for cell in row])
            print(col_str)
        print()

    @classmethod
    def get_float_array(cls, sheet: XSpreadsheet, range_name: str) -> List[List[float]]:
        """
        Gets a 2-Dimensional List of floats

        Args:
            sheet (XSpreadsheet): Spreadsheet to get the float values from.
            range_name (str): Range to get array of floats frm such as 'A1:E18'

        Returns:
            List[List[float]]: 2-Dimensional List of floats
        """
        return cls._convert_to_floats_2d(cls.get_array(sheet=sheet, range_name=range_name))

    get_doubles_array = get_float_array

    # region    convert_to_floats()

    @classmethod
    def _convert_to_floats_1d(cls, vals: Sequence[object]) -> List[float]:
        doubles = []
        for val in vals:
            doubles.append(cls.convert_to_float(val))
        return doubles

    @classmethod
    def _convert_to_floats_2d(cls, vals: Sequence[Sequence[object]]) -> List[List[float]]:
        row_len = len(vals)
        if row_len == 0:
            return []
        col_len = len(vals[0])

        doubles = TableHelper.make_2d_array(num_rows=row_len, num_cols=col_len)
        for row in range(row_len):
            for col in range(col_len):
                doubles[row][col] = cls.convert_to_float(vals[row][col])
        return doubles

    @overload
    @staticmethod
    def convert_to_floats(vals: Sequence[object]) -> List[float]:
        """
        Converts a 1-Dimensional array into List of float

        Args:
            vals (Sequence[object]): List to convert to floats.

        Returns:
            List[float]: vals converted to float
        """
        ...

    @overload
    @staticmethod
    def convert_to_floats(vals: Sequence[Sequence[object]]) -> List[List[float]]:
        """
        Converts a 2-Dimensional array into List of float

        Args:
            vals (Sequence[Sequence[object]]): 2-Dimensional list to convert to floats

        Returns:
            List[List[float]]: 2-Dimensional list of floats.
        """
        ...

    @classmethod
    def convert_to_floats(cls, vals):
        v_len = len(vals)
        if v_len == 0:
            return []
        first = vals[0]
        if GenUtil.is_iterable(arg=first):
            return cls._convert_to_floats_2d(vals)
        else:
            return cls._convert_to_floats_1d(vals)

    convert_to_doubles = convert_to_floats

    # endregion convert_to_floats()

    # endregion ------------- set/get values in 2D array --------------

    # region --------------- set/get rows and columns ------------------

    # region    set_col()
    @overload
    @staticmethod
    def set_col(sheet: XSpreadsheet, values: Sequence[Any], cell_name: str) -> None:
        """
        Inserts a colum of data into spreadsheet

        Args:
            sheet (XSpreadsheet): Spreadsheet
            values (Sequence[Any]): Column Data
            cell_name (str): Name of Cell to begin the insert such as 'A1'
        """
        ...

    @overload
    @staticmethod
    def set_col(sheet: XSpreadsheet, values: Sequence[Any], col_start: int, row_start: int) -> None:
        """
        Inserts a colum of data into spreadsheet

        Args:
            sheet (XSpreadsheet): Spreadsheet
            values (Sequence[Any]):  Column Data as 1-Dimensional Sequence such as a list of values
            col_start (int): Zero-base column index
            row_start (int): Zero-base row index
        """
        ...

    @classmethod
    def set_col(cls, *args, **kwargs) -> None:
        ordered_keys = (1, 2, 3, 4)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('sheet', 'values','cell_name', 'col_start', 'row_start')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("set_col() got an unexpected keyword argument")
            ka[1] = kwargs.get("sheet", None)
            ka[2] = kwargs.get("values", None)
            keys = ("cell_name", "col_start")
            for key in keys:
                if key in kwargs:
                    ka[3] = kwargs[key]
                    break
            if count == 3:
                return ka
            ka[4] = kwargs.get("row_start", None)
            return ka

        if not count in (3, 4):
            raise TypeError("set_col() got an invalid numer of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 3:
            pos = cls.get_cell_position(kargs[3])
            x = pos.X
            y = pos.Y
        else:
            x = kargs[3]
            y = kargs[4]
        values = cast(Sequence[Any], kargs[2])
        val_len = len(values)  # values

        cell_range = cls.get_cell_range(sheet=kargs[1], start_col=x, start_row=y, end_col=x, end_row=y + val_len - 1)
        xcell: XCell = None
        for val in range(val_len):
            xcell = cls.get_cell(cell_range=cell_range, column=0, row=val)
            cls.set_val(cell=xcell, value=values[val])

    # endregion set_col()

    # region    set_row()
    @overload
    @staticmethod
    def set_row(sheet: XSpreadsheet, values: Sequence[Any], cell_name: str) -> None:
        """
        Inserts a row of data into spreadsheet

        Args:
            sheet (XSpreadsheet): Spreadsheet
            values (Sequence[Any]): Row Data
            cell_name (str): Name of Cell to begin the insert such as 'A1'
        """
        ...

    @overload
    @staticmethod
    def set_row(sheet: XSpreadsheet, values: Sequence[Any], col_start: int, row_start: int) -> None:
        """
        Inserts a row of data into spreadsheet

        Args:
            sheet (XSpreadsheet): Spreadsheet
            values (Sequence[Any]):  Row Data as 1-Dimensional Sequence such as a list of values
            col_start (int): Zero-base column index
            row_start (int): Zero-base row index
        """
        ...

    @classmethod
    def set_row(cls, *args, **kwargs) -> None:
        ordered_keys = (1, 2, 3, 4)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('sheet', 'values', 'cell_name', 'col_start', 'row_start')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("set_row() got an unexpected keyword argument")
            ka[1] = kwargs.get("sheet", None)
            ka[2] = kwargs.get("values", None)
            keys = ("cell_name", "col_start")
            for key in keys:
                if key in kwargs:
                    ka[3] = kwargs[key]
                    break
            if count == 3:
                return ka
            ka[4] = kwargs.get("row_start", None)
            return ka

        if not count in (3, 4):
            raise TypeError("set_row() got an invalid numer of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 3:
            pos = cls.get_cell_position(kargs[3])
            col_start = pos.X
            row_start = pos.Y
        else:
            col_start = kargs[3]
            row_start = kargs[4]

        values = cast(Sequence[Any], kargs[2])
        cell_range = cls._get_cell_range_col_row(
            sheet=kargs[1],
            start_col=col_start,
            start_row=row_start,
            end_col=col_start + len(values) - 1,
            end_row=row_start,
        )
        cr_data = mLo.Lo.qi(XCellRangeData, cell_range)
        cr_data.setDataArray(TableHelper.to_2d_tuple(values))  #  1-row 2D array

    # endregion set_row()

    @classmethod
    def get_row(cls, sheet: XSpreadsheet, range_name: str) -> Sequence[Any] | None:
        """
        Gets a row of data from spreadsheet

        Args:
            sheet (XSpreadsheet): Spreadsheet
            range_name (str): Range such as 'A1:A12'

        Returns:
            List[Any] | None: 1-Dimensional List of values on success; Otherwise, None
        """
        vals = cls.get_array(sheet=sheet, range_name=range_name)
        return cls.extract_row(vals=vals, row_idx=0)

    @staticmethod
    def extract_row(vals: Sequence[Sequence[Any]], row_idx: int) -> Sequence[Any] | None:
        row_len = len(vals)
        if row_idx < 0 or row_idx > row_len - 1:
            print("Row index out of range")
            return None

        return vals[row_idx]

    @classmethod
    def get_col(cls, sheet: XSpreadsheet, range_name: str) -> List[Any] | None:
        """
        Gets a column of data from spreadsheet

        Args:
            sheet (XSpreadsheet): Spreadsheet
            range_name (str): Range such as 'A1:A12'

        Returns:
            List[Any] | None: 1-Dimensional List of values on success; Otherwise, None
        """
        vals = cls.get_array(sheet=sheet, range_name=range_name)
        return cls.extract_col(vals=vals, col_idx=0)

    @staticmethod
    def extract_col(vals: Sequence[Sequence[Any]], col_idx: int) -> List[Any] | None:
        row_len = len(vals)
        if row_len == 0:
            return None
        col_len = len(vals[0])
        if col_idx < 0 or col_idx > col_len - 1:
            print("Column index out of range")
            return None

        col_vals = []
        for row in vals:
            col_vals.append(row[col_idx])
        return col_vals

    # endregion --------------- set/get rows and columns -----------------

    # region --------------- special cell types ------------------------

    @classmethod
    def set_date(cls, sheet: XSpreadsheet, cell_name: str, day: int, month: int, year: int) -> None:
        """Writes a date with standard date format into a spreadsheet"""
        xcell = cls._get_cell_sheet_cell(sheet=sheet, cell_name=cell_name)
        xcell.setFormula(f"{month}/{day}/{year}")

        nfs_supplier = mLo.Lo.create_instance_mcf(XNumberFormatsSupplier, "com.sun.star.util.NumberFormatsSupplier")
        if nfs_supplier is None:
            return
        number_formats = nfs_supplier.getNumberFormats()
        xformat_types = mLo.Lo.qi(XNumberFormatTypes, number_formats)
        if xformat_types is None:
            return
        alocale = Locale()
        # aLocale.Country = "GB"
        # aLocale.Language = "en"

        nformat = xformat_types.getStandardFormat(NumberFormat.DATE, alocale)
        mProps.Props.set_property(prop_set=xcell, name="NumberFormat", value=nformat)

    # region    add_annotation()
    @overload
    @staticmethod
    def add_annotation(sheet: XSpreadsheet, cell_name: str, msg: str) -> XSheetAnnotation | None:
        """
        Adds an annotation to a cell and makes the annotation visible.

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cell_name (str): Name of cell to add annotation such as 'A1'
            msg (str): Annotation Text

        Returns:
            XSheetAnnotation | None: Cell annotation that was addded on success; Othwrwise, None
        """
        ...

    @overload
    @staticmethod
    def add_annotation(sheet: XSpreadsheet, cell_name: str, msg: str, is_visible: bool) -> XSheetAnnotation | None:
        """
        Adds an annotation to a cell

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cell_name (str): Name of cell to add annotation such as 'A1'
            msg (str): Annotation Text
            set_visible (bool): Determines if the annaction is set visible

        Returns:
            XSheetAnnotation | None: Cell annotation that was addded on success; Othwrwise, None
        """
        ...

    @classmethod
    def add_annotation(cls, sheet: XSpreadsheet, cell_name: str, msg: str, is_visible=True) -> XSheetAnnotation | None:
        # add the annotation
        addr = cls.get_cell_address(sheet=sheet, cell_name=cell_name)
        anns_supp = mLo.Lo.qi(XSheetAnnotationsSupplier, sheet)
        if anns_supp is None:
            return
        anns = anns_supp.getAnnotations()
        anns.insertNew(addr, msg)

        # get a reference to the annotation
        xcell = cls.get_cell(sheet=sheet, cell_name=cell_name)
        ann_anchor = mLo.Lo.qi(XSheetAnnotationAnchor, xcell)
        ann = ann_anchor.getAnnotation()
        ann.setIsVisible(is_visible)
        return ann

    # endregion add_annotation()

    @classmethod
    def get_annotation(cls, sheet: XSpreadsheet, cell_name: str) -> XSheetAnnotation | None:
        """
        Gets an annotation of a cell

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cell_name (str): Name of cell to add annotation such as 'A1'

        Returns:
            XSheetAnnotation | None: Cell annotation on success; Othwrwise, None
        """
        # get a reference to the annotation
        xcell = cls.get_cell(sheet=sheet, cell_name=cell_name)
        ann_anchor = mLo.Lo.qi(XSheetAnnotationAnchor, xcell)
        if ann_anchor is None:
            print(f"No XSheetAnnotationAnchor for {cell_name}")
            return None
        return ann_anchor.getAnnotation()

    @classmethod
    def get_annotation_str(cls, sheet: XSpreadsheet, cell_name: str) -> str:
        """
        Gets text of an annotation for a cell.

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cell_name (str): Name of cell to add annotation such as 'A1'

        Returns:
            str: Cell annotation text
        """
        ann = cls.get_annotation(sheet=sheet, cell_name=cell_name)
        if ann is None:
            return ""
        xsimple_text = mLo.Lo.qi(XSimpleText, ann)
        if xsimple_text is None:
            return ""
        return xsimple_text.getString()

    # endregion ------------ special cell types ------------------------

    # region --------------- get XCell and XCellRange methods ----------

    # region    get_cell()
    @classmethod
    def _get_cell_sheet_col_row(cls, sheet: XSpreadsheet, column: int, row: int) -> XCell | None:
        return sheet.getCellByPosition(column, row)

    @classmethod
    def _get_cell_sheet_addr(cls, sheet: XSpreadsheet, addr: CellAddress) -> XCell | None:
        # not using Sheet value in addr
        return cls._get_cell_sheet_col_row(sheet=sheet, column=addr.Column, row=addr.Row)

    @classmethod
    def _get_cell_sheet_cell(cls, sheet: XSpreadsheet, cell_name: str) -> XCell | None:
        cell_range = sheet.getCellRangeByName(cell_name)
        return cls._get_cell_cell_rng(cell_range=cell_range, column=0, row=0)

    @classmethod
    def _get_cell_cell_rng(cls, cell_range: XCellRange, column: int, row: int) -> XCell | None:
        return cell_range.getCellByPosition(column, row)

    @overload
    @staticmethod
    def get_cell(sheet: XSpreadsheet, addr: CellAddress) -> XCell | None:
        ...

    @overload
    @staticmethod
    def get_cell(sheet: XSpreadsheet, cell_name: str) -> XCell | None:
        ...

    @overload
    @staticmethod
    def get_cell(sheet: XSpreadsheet, column: int, row: int) -> XCell | None:
        ...

    @overload
    @staticmethod
    def get_cell(cell_range: XCellRange) -> XCell | None:
        ...

    @overload
    @staticmethod
    def get_cell(cell_range: XCellRange, column: int, row: int) -> XCell | None:
        ...

    @classmethod
    def get_cell(cls, *args, **kwargs) -> XCell | None:
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('sheet', 'cell_range','addr', 'column', 'cell_name', 'row')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_cell() got an unexpected keyword argument")
            keys = ("sheet", "cell_range")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            keys = ("addr", "column", "cell_name")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("row", None)
            return ka

        if not count in (1, 2, 3):
            raise TypeError("get_cell() got an invalid numer of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            # get_cell(cell_range: XCellRange)
            # cell range is relative position.
            # if a range is C4:E9 then Cell range at col=0 ,row=0 is C4
            return cls._get_cell_cell_rng(cell_range=kargs[1], column=0, row=0)

        elif count == 2:
            if isinstance(kargs[2], str):
                # get_cell(sheet: XSpreadsheet, cell_name: str)
                return cls._get_cell_sheet_cell(sheet=kargs[1], cell_name=kargs[2])
            else:
                # get_cell(sheet: XSpreadsheet, addr: CellAddress)
                return cls._get_cell_sheet_addr(sheet=kargs[1], addr=kargs[2])
        else:
            sheet = mLo.Lo.qi(XSpreadsheet, kargs[1])
            if sheet is None:
                # get_cell(cell_range: XCellRange, column: int, row: int)
                return cls._get_cell_cell_rng(cell_range=kargs[1], column=kargs[2], row=kargs[3])
            else:
                # get_cell(sheet: XSpreadsheet, column: int, row: int)
                return cls._get_cell_sheet_col_row(sheet=sheet, column=kargs[2], row=kargs[3])

    # endregion get_cell()

    @staticmethod
    def is_cell_range_name(s: str) -> bool:
        return ":" in s

    @staticmethod
    def is_single_cell_range(cr_addr: CellRangeAddress) -> bool:
        return cr_addr.StartColumn == cr_addr.EndColumn and cr_addr.StartRow == cr_addr.EndRow

    # region    get_cell_range()
    @classmethod
    def _get_cell_range_addr(cls, sheet: XSpreadsheet, addr: CellRangeAddress) -> XCellRange | None:
        return cls._get_cell_range_col_row(
            sheet=sheet,
            start_col=addr.StartColumn,
            start_row=addr.StartRow,
            end_col=addr.EndColumn,
            end_row=addr.EndRow,
        )

    @staticmethod
    def _get_cell_range_rng_name(sheet: XSpreadsheet, range_name: str) -> XCellRange | None:
        cell_range = sheet.getCellRangeByName(range_name)
        if cell_range is None:
            print(f"Could not access cell range: '{range_name}'")
            return None
        return cell_range

    @staticmethod
    def _get_cell_range_col_row(
        sheet: XSpreadsheet, start_col: int, start_row: int, end_col: int, end_row: int
    ) -> XCellRange | None:
        try:
            return sheet.getCellRangeByPosition(start_col, start_row, end_col, end_row)
        except Exception:
            print(f"Could not access cell range : ({start_col}, {start_row}, {end_col}, {end_row})")
        return None

    @overload
    @staticmethod
    def get_cell_range(sheet: XSpreadsheet, cr_addr: CellRangeAddress) -> XCellRange | None:
        """
        Gets a cell range

        Args:
            sheet (XSpreadsheet): Spreadsheet Document
            addr (CellRangeAddress): Cell range Address

        Returns:
            XCellRange | None: Cell range is successful; Otherwise, None
        """
        ...

    @overload
    @staticmethod
    def get_cell_range(sheet: XSpreadsheet, range_name: str) -> XCellRange | None:
        """
        Gets a cell range

        Args:
            sheet (XSpreadsheet): Spreadsheet Document
            range_name (str): Range Name such as 'A1:D5'

        Returns:
            XCellRange | None: Cell range is successful; Otherwise, None
        """
        ...

    @overload
    @staticmethod
    def get_cell_range(
        sheet: XSpreadsheet, start_col: int, start_row: int, end_col: int, end_row: int
    ) -> XCellRange | None:
        """
        Gets a cell range

        Args:
            sheet (XSpreadsheet): Spreadsheet Document
            start_col (int): Start Column
            start_row (int): Start Row
            end_col (int): End Column
            end_row (int): End Row

        Returns:
            XCellRange | None: Cell range is successful; Otherwise, None
        """
        ...

    @classmethod
    def get_cell_range(cls, *args, **kwargs) -> XCellRange | None:
        ordered_keys = (1, 2, 3, 4, 5)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('sheet', 'cr_addr','range_name', 'start_col', 'start_row', 'end_col', 'end_row')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_cell_range() got an unexpected keyword argument")
            ka[1] = kwargs.get("sheet", None)
            keys = ("cr_addr", "range_name", "start_col")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("start_row", None)
            ka[4] = kwargs.get("end_col", None)
            ka[5] = kwargs.get("end_row", None)
            return ka

        if not count in (2, 5):
            raise TypeError("get_cell_range() got an invalid numer of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 2:
            if isinstance(kargs[2], str):
                # def get_cell_range(sheet: XSpreadsheet, range_name: str)
                return cls._get_cell_range_rng_name(sheet=kargs[1], range_name=kargs[2])
            else:
                # get_cell_range(sheet: XSpreadsheet, addr:CellRangeAddress)
                return cls._get_cell_range_addr(sheet=kargs[1], addr=kargs[2])
        else:
            # get_cell_range(sheet: XSpreadsheet, start_col: int, start_row: int, end_col: int, end_row: int)
            return cls._get_cell_range_col_row(
                sheet=kargs[1],
                start_col=kargs[2],
                start_row=kargs[3],
                end_col=kargs[4],
                end_row=kargs[5],
            )

    # endregion get_cell_range()

    # region    find_used_range()

    @overload
    @staticmethod
    def find_used_range(sheet: XSpreadsheet) -> XCellRange | None:
        """
        Find used range

        Args:
            sheet (XSpreadsheet): Spreadsheet Document

        Returns:
            XCellRange | None: Cell range on success; Othwrwise, None
        """
        ...

    @overload
    @staticmethod
    def find_used_range(sheet: XSpreadsheet, cell_name: str) -> XCellRange | None:
        """
        Find used range

        Args:
            sheet (XSpreadsheet): Spreadsheet Document
            cell_name (str): Cell Name

        Returns:
            XCellRange | None: Cell range on success; Othwrwise, None
        """
        ...

    @classmethod
    def find_used_range(cls, sheet: XSpreadsheet, cell_name: str = None) -> XCellRange | None:
        if cell_name is None:
            cursor = sheet.createCursor()
        else:
            xrange = cls._get_cell_range_rng_name(sheet=sheet, range_name=cell_name)
            cell_range = mLo.Lo.qi(XSheetCellRange, xrange)
            cursor = sheet.createCursorByRange(cell_range)
        return cls.find_used_cursor(cursor)

    # endregion find_used_range()

    @staticmethod
    def find_used_cursor(cursor: XSheetCellCursor) -> XCellRange | None:
        """
        Find used cursor

        Args:
            cursor (XSheetCellCursor): Sheet Cursor

        Returns:
            XCellRange | None: Cell range on success; Othwrwise, None
        """
        # find the used area
        ua_cursor = mLo.Lo.qi(XUsedAreaCursor, cursor)
        if ua_cursor is None:
            return None
        ua_cursor.gotoStartOfUsedArea(False)
        ua_cursor.gotoEndOfUsedArea(True)

        used_range = mLo.Lo.qi(XCellRange, ua_cursor)
        return used_range

    @staticmethod
    def get_col_range(sheet: XSpreadsheet, idx: int) -> XCellRange | None:
        """
        Get Column by index

        Args:
            sheet (XSpreadsheet): Spreadsheet
            idx (int): Zero-based column index

        Returns:
            XCellRange | None: Cell range on success; Othwrwise, None
        """
        cr_range = mLo.Lo.qi(XColumnRowRange, sheet)
        if cr_range is None:
            return None
        cols = cr_range.getColumns()
        con = mLo.Lo.qi(XIndexAccess, cols)
        try:
            return mLo.Lo.qi(XCellRange, con.getByIndex(idx))
        except Exception:
            print(f"Could not access range for column position: {idx}")
        return None

    @staticmethod
    def get_row_range(sheet: XSpreadsheet, idx: int) -> XCellRange | None:
        """
        Get Row by index

        Args:
            sheet (XSpreadsheet): Spreadsheet
            idx (int): Zero-based column index

        Returns:
            XCellRange | None: Cell range on success; Othwrwise, None
        """
        cr_range = mLo.Lo.qi(XColumnRowRange, sheet)
        if cr_range is None:
            return None
        rows = cr_range.getRows()
        con = con = mLo.Lo.qi(XIndexAccess, rows)
        try:
            return mLo.Lo.qi(XCellRange, con.getByIndex(idx))
        except Exception:
            print(f"Could not access range for row position: {idx}")
        return None

    # endregion ------------ get XCell and XCellRange methods ----------

    # region --------------- convert cell/cellrange names to positions -

    @classmethod
    def get_cell_range_positions(cls, range_name: str) -> Tuple[Point, Point] | None:
        """
        Gets Cell range as a tuple of Point, Point

        First Point.X is start column index, Point.Y is start row index.
        Second Point.X is end column index, Point.Y is end row index.

        Args:
            range_name (str): Range name such as 'A1:C8'

        Returns:
            Tuple[Point, Point] | None: Range as tuple on success; Otherwise, None
        """
        cell_names = range_name.split(":")
        if len(cell_names) != 2:
            print(f"Cell range not found in {range_name}")
            return None
        start_pos = cls.get_cell_position(cell_names[0])
        end_pos = cls.get_cell_position(cell_names[1])
        return (start_pos, end_pos)

    # region    get_cell_position()
    @classmethod
    def get_cell_position(cls, cell_name: str) -> Point | None:
        """
        Gets a cell name as a Point.

        Point.X is column index.
        Point.y is row index.

        Args:
            cell_name (str): Cell name suca as 'A1'

        Returns:
            Point | None: cell name as Point on success; Otherwise, None
        """
        m = cls._rx_cell.match(cell_name)
        if m:
            ncolumn = cls.column_string_to_number(str(m.group(1)).upper())
            nrow = cls.row_string_to_number(m.group(2))
            return Point(ncolumn, nrow)
        else:
            print("No match found")
        return None

    @classmethod
    def get_cell_pos(cls, sheet: XSpreadsheet, cell_name: str) -> Point:
        """
        Contains the position of the top left cell of this range in the sheet (in 1/100 mm).

        This property contains the absolute position in the whole sheet,
        not the position in the visible area.

        Args:
            cell_name (str):  Cell name suca as 'A1'
            sheet (XSpreadsheet): Spreadsheet

        Returns:
            Point | None: cell name as Point on success; Otherwise, None
        """
        xcell = cls._get_cell_sheet_cell(sheet=sheet, cell_name=cell_name)
        if xcell is None:
            print("No match found")
            return None
        pos = mProps.Props.get_property(xprops=xcell, name="Position")
        if pos is None:
            print(f"Could not determine position of cell '{cell_name}'")
            pos = cls.CELL_POS
            # print("No match found")
            # return None
        return pos

    # endregion get_cell_position()

    @staticmethod
    def column_string_to_number(col_str: str) -> int:
        """
        Converts a Column Name into an int.
        Results are zero based so ``a`` converts to ``0``

        Args:
            col_str (str):Case insensitive column name such as 'a' or 'AB'

        Returns:
            int: Zero based int representing column name
        """
        i = TableHelper.col_name_to_int(name=col_str)
        return i - 1  # convert to zero based.

    @staticmethod
    def row_string_to_number(row_str: str) -> int:
        """
        Converts a string containing an int into an int

        Args:
            row_str (str): string to convert

        Returns:
            int: Number if conversion succeeds; Othwrwise, 0
        """
        try:
            return TableHelper.row_name_to_int(row_str) - 1
        except ValueError:
            print(f"Incorrect format for {row_str}")
        return 0

    # endregion ----------- convert cell/cellrange names to positions --

    # region --------------- get cell and cell range addresses ---------

    # region    get_cell_address()

    @staticmethod
    def _get_cell_address_cell(cell: XCell) -> CellAddress | None:
        addr = mLo.Lo.qi(XCellAddressable, cell)
        if addr is None:
            return None
        return addr.getCellAddress()

    @classmethod
    def _get_cell_address_sheet(cls, sheet: XSpreadsheet, cell_name: str) -> CellAddress | None:
        cell_range = sheet.getCellRangeByName(cell_name)
        start_cell = cls._get_cell_cell_rng(cell_range=cell_range, column=0, row=0)
        return cls._get_cell_address_cell(start_cell)

    @overload
    @staticmethod
    def get_cell_address(cell: XCell) -> CellAddress | None:
        """
        Gets Cell Address

        Args:
            cell (XCell): Cell

        Returns:
            CellAddress | None: Cell Address on success; Othwrwise, None
        """
        ...

    @overload
    @staticmethod
    def get_cell_address(sheet: XSpreadsheet, cell_name: str) -> CellAddress | None:
        """
        Gets Cell Address

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cell_name (str): Cell name such as 'A1'

        Returns:
            CellAddress | None: Cell Address on success; Othwrwise, None
        """
        ...

    @overload
    @staticmethod
    def get_cell_address(sheet: XSpreadsheet, addr: CellAddress) -> CellAddress | None:
        """
        Gets Cell Address

        Args:
            sheet (XSpreadsheet): Spreadsheet
            addr (CellAddress): Cell Address

        Returns:
            CellAddress | None: Cell Address on success; Othwrwise, None
        """
        ...

    @overload
    @staticmethod
    def get_cell_address(sheet: XSpreadsheet, col: int, row: int) -> CellAddress | None:
        """
        Gets Cell Address

        Args:
            sheet (XSpreadsheet): Spreadsheet
            col (int): Zero-base column index
            row (int): Zero-base row index

        Returns:
            CellAddress | None: Cell Address on success; Othwrwise, None
        """
        ...

    @classmethod
    def get_cell_address(cls, *args, **kwargs) -> CellAddress | None:
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('cell', 'sheet','cell_name', 'col', 'addr', 'row')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_cell_address() got an unexpected keyword argument")
            keys = ("cell", "sheet")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            keys = ("cell_name", "col", "addr")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("row", None)
            return ka

        if not count in (1, 2, 3):
            raise TypeError("get_cell_address() got an invalid numer of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            return cls._get_cell_address_cell(cell=kargs[1])
        elif count == 2:
            if isinstance(kargs[2], str):
                return cls._get_cell_address_sheet(sheet=kargs[1], cell_name=kargs[2])
            cell_name = cls._get_cell_str_addr(addr=kargs[2])
            return cls._get_cell_address_sheet(sheet=kargs[1], cell_name=cell_name)
        elif count == 3:
            cell_name = cls._get_cell_str_col_row(col=kargs[2], row=kargs[3])
            return cls._get_cell_address_sheet(sheet=kargs[1], cell_name=cell_name)

    # endregion get_cell_address()

    # region    get_address()
    @staticmethod
    def _get_address_cell(cell_range: XCellRange) -> CellRangeAddress | None:
        addr = mLo.Lo.qi(XCellRangeAddressable, cell_range)
        if addr is None:
            return None
        return addr.getRangeAddress()

    @classmethod
    def _get_address_sht_rng(cls, sheet: XSpreadsheet, range_name: str) -> CellRangeAddress | None:
        return cls._get_address_cell(cls._get_cell_range_rng_name(sheet=sheet, range_name=range_name))

    @overload
    @staticmethod
    def get_address(cell_range: XCellRange) -> CellRangeAddress | None:
        """
        Gets Range Address

        Args:
            cell_range (XCellRange): Cell Range

        Returns:
            CellRangeAddress | None: Cell Range Address on success; Othwrwise, None
        """
        ...

    @overload
    @staticmethod
    def get_address(sheet: XSpreadsheet, range_name: str) -> CellRangeAddress | None:
        """
        Gets Range Address

        Args:
            sheet (XSpreadsheet): Spreadsheet
            range_name (str): Range name such as 'A1:D7'

        Returns:
            CellRangeAddress | None: Cell Range Address on success; Othwrwise, None
        """
        ...

    @overload
    @staticmethod
    def get_address(
        sheet: XSpreadsheet, start_col: int, start_row: int, end_col: int, end_row: int
    ) -> CellRangeAddress | None:
        """
        Gets Range Address

        Args:
            sheet (XSpreadsheet): Spreadsheet
            start_col (int): Zero-base start column index
            start_row (int): Zero-base start row index
            end_col (int): Zero-base end column index
            end_row (int): Zero-base end row index

        Returns:
            CellRangeAddress | None: Cell Range Address on success; Othwrwise, None
        """
        ...

    @classmethod
    def get_address(cls, *args, **kwargs) -> CellRangeAddress | None:
        ordered_keys = (1, 2, 3, 4, 5)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('cell_range', 'sheet', 'range_name', 'start_col', 'start_row', 'end_col' , 'end_row')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_address() got an unexpected keyword argument")
            keys = ("cell_range", "sheet")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            keys = ("range_name", "start_col")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("start_row", None)
            ka[4] = kwargs.get("end_col", None)
            ka[5] = kwargs.get("end_row", None)
            return ka

        if not count in (1, 2, 5):
            raise TypeError("get_address() got an invalid numer of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            return cls._get_address_cell(cell_range=kargs[1])
        elif count == 2:

            return cls._get_address_sht_rng(sheet=kargs[1], range_name=kargs[2])
        else:
            range_name = cls._get_range_str_col_row(
                start_col=kargs[2], start_row=kargs[3], end_col=kargs[4], end_row=kargs[5]
            )
            return cls._get_address_sht_rng(sheet=kargs[1], range_name=range_name)

    # endregion get_address()

    # region    print_cell_address()
    @overload
    @staticmethod
    def print_cell_address(cell: XCell) -> None:
        """
        Prints Cell to terminal such as ``Cell: Sheet1.D3``

        Args:
            cell (XCell): cell
        """
        ...

    @overload
    @staticmethod
    def print_cell_address(addr: CellAddress) -> None:
        """
        Prints Cell to terminal such as ``Cell: Sheet1.D3``

         Args:
             addr (CellAddress): Cell Address
        """
        ...

    @classmethod
    def print_cell_address(cls, *args, **kwargs) -> None:
        ordered_keys = (1,)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('cell', 'addr')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("print_cell_address() got an unexpected keyword argument")
            keys = ("cell", "addr")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            return ka

        if count != 1:
            raise TypeError("print_cell_address() got an invalid numer of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if mInfo.Info.is_type_interface(obj=kargs[1], type_name="com.sun.star.table.XCell"):
            addr = cls._get_cell_address_cell(cell=kargs[1])
        else:
            addr = kargs[1]
        print(f"Cell: Sheet{addr.Sheet+1}.{cls.get_cell_str(addr=addr)}")

    # endregion    print_cell_address()

    # region    print_address()
    @overload
    @staticmethod
    def print_address(cell_range: XCellRange) -> None:
        """
        Prints Cell range to terminal such as ``'Range: Sheet1.C3:F22``

        Args:
            cell_range (XCellRange): Cell range
        """
        ...

    @overload
    @staticmethod
    def print_address(cr_addr: CellRangeAddress) -> None:
        """
        Prints Cell range to terminal such as ``'Range: Sheet1.C3:F22``

        Args:
            cr_addr (CellRangeAddress): Cell Address
        """
        ...

    @classmethod
    def print_address(cls, *args, **kwargs) -> None:
        ordered_keys = (1,)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('cell_range', 'cr_addr')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("print_address() got an unexpected keyword argument")
            keys = ("cell_range", "cr_addr")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            return ka

        if count != 1:
            raise TypeError("print_address() got an invalid numer of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        cell_range = mLo.Lo.qi(XCellRange, kargs[1])
        if cell_range is None:
            # when cast is used with an import the is not available at runtime must be quoted.
            cr_addr = cast("CellRangeAddress", kargs[1])
        else:
            cr_addr = cls._get_address_cell(cell_range=kargs[1])
        msg = f"Range: Sheet{cr_addr.Sheet+1}.{cls.get_cell_str(col=cr_addr.StartColumn,row=cr_addr.StartRow)}:"
        msg += f"{cls.get_cell_str(col=cr_addr.EndColumn, row=cr_addr.EndRow)}"
        print(msg)

    # endregion  print_address()

    @classmethod
    def print_addresses(cls, *cr_addrs: CellRangeAddress) -> None:
        """
        Prints Address for one or more CellRangeAddress
        """
        print(f"No of cellrange addresses: {len(cr_addrs)}")
        for cr_addr in cr_addrs:
            cls.print_address(cr_addr=cr_addr)
        print()

    @staticmethod
    def get_cell_series(sheet: XSpreadsheet, range_name: str) -> XCellSeries | None:
        """
        Get cell series for a range

        Args:
            sheet (XSpreadsheet): Spreadsheet
            range_name (str): Range name such as 'A1:B7'

        Returns:
            XCellSeries | None: Cell series on success; Otherwise, None
        """
        cell_range = sheet.getCellRangeByName(range_name)
        return mLo.Lo.qi(XCellSeries, cell_range)

    # region    is_equal_addresses()

    @overload
    @staticmethod
    def is_equal_addresses(addr1: CellRangeAddress, addr2: CellRangeAddress) -> bool:
        """
        Gets if two instances of CellRangeAddress are equal

        Args:
            addr1 (CellRangeAddress): Cell Range Address
            addr2 (CellRangeAddress): Cell Range Address

        Returns:
            bool: True if equal; Otherwise, False
        """
        ...

    @overload
    @staticmethod
    def is_equal_addresses(addr1: CellAddress, addr2: CellAddress) -> bool:
        """
        Gets if two instances of CellAddress are equal

        Args:
            addr1 (CellAddress): Cell Address
            addr2 (CellAddress): Cell Address

        Returns:
            bool: True if equal; Otherwise, False
        """
        ...

    @staticmethod
    def is_equal_addresses(addr1: object, addr2: object) -> bool:
        if addr1 is None or addr2 is None:
            return False
        try:
            is_same_type = addr1.typeName == addr2.typeName
            if not is_same_type:
                return False
        except AttributeError:
            return False
        if mInfo.Info.is_type_struct(addr1, "com.sun.star.table.CellAddress"):
            # when cast is used with an import the is not available at runtime must be quoted.
            a = cast("CellAddress", addr1)
            b = cast("CellAddress", addr2)
            return a.Sheet == b.Sheet and a.Column == b.Column and a.Row == b.Row
        if mInfo.Info.is_type_struct(addr1, "com.sun.star.table.CellRangeAddress"):
            # when cast is used with an import the is not available at runtime must be quoted.
            a = cast("CellRangeAddress", addr1)
            b = cast("CellRangeAddress", addr2)
            return (
                a.Sheet == b.Sheet
                and a.StartColumn == b.StartColumn
                and a.StartRow == b.StartRow
                and a.EndColumn == b.EndColumn
                and a.EndRow == b.EndRow
            )
        return False

    # endregion  is_equal_addresses()

    # endregion ------------ get cell and cell range addresses ---------

    # region --------------- convert cell range address to string ------

    # region    get_range_str()
    @classmethod
    def _get_range_str_cell_rng_sht(cls, cell_range: XCellRange, sheet: XSpreadsheet) -> str:
        """return as str using the name taken from the sheet works, Sheet1.A1:B2"""
        return cls._get_range_str_cr_addr_sht(cls._get_address_cell(cell_range=cell_range), sheet)

    @classmethod
    def _get_range_str_cr_addr_sht(cls, cr_addr: CellRangeAddress, sheet: XSpreadsheet) -> str:
        """return as str using the name taken from the sheet works, Sheet1.A1:B2"""
        return f"{cls.get_sheet_name(sheet=sheet)}.{cls._get_range_str_cr_addr(cr_addr)}"

    @classmethod
    def _get_range_str_cell_rng(cls, cell_range: XCellRange) -> str:
        """return as str, A1:B2"""
        return cls._get_range_str_cr_addr(cls._get_address_cell(cell_range=cell_range))

    @classmethod
    def _get_range_str_cr_addr(cls, cr_addr: CellRangeAddress) -> str:
        """return as str, A1:B2"""
        result = f"{cls._get_cell_str_col_row(cr_addr.StartColumn, cr_addr.StartRow)}:"
        result += f"{cls._get_cell_str_col_row(cr_addr.EndColumn, cr_addr.EndRow)}"
        return result

    @classmethod
    def _get_range_str_col_row(cls, start_col: int, start_row: int, end_col: int, end_row: int) -> str:
        """return as str, A1:B2"""
        return f"{cls._get_cell_str_col_row(start_col, start_row)}:{cls._get_cell_str_col_row(end_col, end_row)}"

    @overload
    @staticmethod
    def get_range_str(cell_range: XCellRange) -> str:
        """
        Gets the range as a string inf format of ``A1:B2``

        Args:
            cell_range (XCellRange): Cell Range
            sheet (XSpreadsheet): Spreadsheet

        Returns:
            str: range a string
        """
        ...

    @overload
    @staticmethod
    def get_range_str(cr_addr: CellRangeAddress) -> str:
        """
        Gets the range as a string inf format of ``A1:B2``

        Args:
            cr_addr (CellRangeAddress): Cell Range Address
            sheet (XSpreadsheet): Spreadsheet

        Returns:
            str: range a string
        """
        ...

    @overload
    @staticmethod
    def get_range_str(cell_range: XCellRange, sheet: XSpreadsheet) -> str:
        """
        Gets the range as a string inf format of ``Sheet1.A1:B2``

        Args:
            cell_range (XCellRange): Cell Range
            sheet (XSpreadsheet): Spreadsheet

        Returns:
            str: range a string
        """
        ...

    @overload
    @staticmethod
    def get_range_str(cr_addr: CellRangeAddress, sheet: XSpreadsheet) -> str:
        """
        Gets the range as a string inf format of ``Sheet1.A1:B2``

        Args:
            cr_addr (CellRangeAddress): Cell Range Address
            sheet (XSpreadsheet): Spreadsheet

        Returns:
            str: range a string
        """
        ...

    @overload
    @staticmethod
    def get_range_str(start_col: int, start_row: int, end_col: int, end_row: int) -> str:
        """
        Gets the range as a string inf format of ``A1:B2``

        Args:
            start_col (int): Zero-based start column index
            start_row (int): Zero-based start row index
            end_col (int): Zero-based end column index
            end_row (int): Zero-based end row index

        Returns:
            str: range a string
        """
        ...

    @overload
    @staticmethod
    def get_range_str(start_col: int, start_row: int, end_col: int, end_row: int, sheet: XSpreadsheet) -> str:
        """
        Gets the range as a string inf format of ``Sheet1.A1:B2``

        Args:
            start_col (int): Zero-based start column index
            start_row (int): Zero-based start row index
            end_col (int): Zero-based end column index
            end_row (int): Zero-based end row index
            sheet (XSpreadsheet): Spreadsheet

        Returns:
            str: range a string
        """
        ...

    @classmethod
    def get_range_str(cls, *args, **kwargs) -> str:
        ordered_keys = (1, 2, 3, 4, 5)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('cell_range', 'cr_addr', 'sheet', 'start_col', 'start_row', 'end_col', 'end_row')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_range_str() got an unexpected keyword argument")
            keys = ("cell_range", "cr_addr", "start_col")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            if count < 5:
                keys = ("sheet", "start_row")
            else:
                keys = ("start_row",)
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("end_col", None)
            ka[4] = kwargs.get("end_row", None)
            if count == 4:
                return ka
            ka[5] = kwargs.get("sheet", None)
            return ka

        if not count in (1, 2, 4, 5):
            raise TypeError("get_range_str() got an invalid numer of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            if mInfo.Info.is_type_interface(kargs[1], "com.sun.star.table.XCellRange"):
                # get_range_str(cell_range: XCellRange)
                return cls._get_range_str_cell_rng(cell_range=kargs[1])
            else:
                # get_range_str(cr_addr: CellRangeAddress)
                return cls._get_range_str_cr_addr(cr_addr=kargs[1])

        elif count == 2:
            if mInfo.Info.is_type_interface(kargs[1], "com.sun.star.table.XCellRange"):
                # def get_range_str(cell_range: XCellRange, sheet: XSpreadsheet)
                return cls._get_range_str_cell_rng_sht(cell_range=kargs[1], sheet=kargs[2])
            else:
                # get_range_str(cr_addr: CellRangeAddress, sheet: XSpreadsheet)
                return cls._get_range_str_cr_addr_sht(cr_addr=kargs[1], sheet=kargs[2])
        elif count == 4:
            # get_range_str(start_col:int, start_row:int, end_col:int, end_row:int)
            return cls._get_range_str_col_row(
                start_col=kargs[1], start_row=kargs[2], end_col=kargs[3], end_row=kargs[4]
            )
        elif count == 5:
            # get_range_str(start_col: int, start_row: int, end_col: int, end_row: int,  sheet: XSpreadsheet)
            rng_str = cls._get_range_str_col_row(
                start_col=kargs[1], start_row=kargs[2], end_col=kargs[3], end_row=kargs[4]
            )
            return f"{cls.get_sheet_name(sheet=kargs[5])}.{rng_str}"
        return ""

    # endregion get_range_str()

    # region    get_cell_str()
    @classmethod
    def _get_cell_str_addr(cls, addr: CellAddress) -> str:
        return cls._get_cell_str_col_row(col=addr.Column, row=addr.Row)

    @classmethod
    def _get_cell_str_col_row(cls, col: int, row: int) -> str:
        if col < 0 or row < 0:
            print("Cell position is negative; using A1")
            return "A1"
        return f"{cls.column_number_str(col)}{row + 1}"

    @classmethod
    def _get_cell_str_cell(cls, cell: XCell) -> str:
        return cls._get_cell_str_addr(cls._get_cell_address_cell(cell=cell))

    @overload
    @staticmethod
    def get_cell_str(addr: CellAddress) -> str:
        """
        Gets the range as a string inf format of ``A1``

        Args:
            addr (CellAddress): Cell Address

        Returns:
            str: Cell as str
        """
        ...

    @overload
    @staticmethod
    def get_cell_str(cell: XCell) -> str:
        """
        Gets the range as a string inf format of ``A1``

        Args:
            cell (XCell): Cell

        Returns:
            str: Cell as str
        """
        ...

    @overload
    @staticmethod
    def get_cell_str(col: int, row: int) -> str:
        """
         Gets the range as a string inf format of ``A1``

        Args:
            col (int): Zero-based column index
            row (int): Zero-based row index

        Returns:
            str: Cell as str
        """
        ...

    @classmethod
    def get_cell_str(cls, *args, **kwargs) -> str:
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('addr', 'cell', 'col', 'row')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("get_cell_str() got an unexpected keyword argument")
            keys = ("addr", "cell", "col")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break

            if count == 1:
                return ka

            ka[2] = kwargs.get("row", None)
            return ka

        if not count in (1, 2):
            raise TypeError("get_cell_str() got an invalid numer of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            # def get_cell_str(addr: CellAddress) or
            # def get_cell_str(cell: XCell)
            if mInfo.Info.is_type_interface(kargs[1], "com.sun.star.table.XCell"):
                return cls._get_cell_str_cell(kargs[1])
            else:
                return cls._get_cell_str_addr(kargs[1])
        else:
            # def get_cell_str(col: int, row: int)
            return cls._get_cell_str_col_row(col=kargs[1], row=kargs[2])

    # endregion get_cell_str()

    @staticmethod
    def column_number_str(col: int) -> str:
        """
        Creates a colum Name from zero base column number.

        Columns are numbered starting at 0 where 0 corresponds to ``A``
        They run as ``A-Z``, ``AA-AZ``, ``BA-BZ``, ..., ``IV``

        Args:
            col (int): Zero based column name

        Returns:
            str: Column Name
        """
        num = col + 1  # shift to one based.
        return TableHelper.make_column_name(num)

    # endregion ------------ convert cell range address to string ------

    # region --------------- search ------------------------------------

    @staticmethod
    def find_all(srch: XSearchable, sd: XSearchDescriptor) -> List[XCellRange] | None:
        """
        Searches spreadsheet and returns a list of Cell Ranges that match search criteria

        Args:
            srch (XSearchable): Searchable object
            sd (XSearchDescriptor): Search descriptro

        Returns:
            List[XCellRange] | None: A list of cell ranges on success; Otherwise, None

        Example:
            .. code-block:: python

                from ooodev.utils.lo import Lo
                from ooodev.office.calc import Calc
                from com.sun.star.util import XSearchable

                doc = Calc.create_doc(loader)
                sheet = Calc.get_sheet(doc=doc, index=0)
                Calc.set_val(value='test', sheet=sheet, cell_name="A1")
                Calc.set_val(value='test', sheet=sheet, cell_name="C3")
                srch = Lo.qi(XSearchable, sheet)
                sd = srch.createSearchDescriptor()
                sd.setSearchString('test')
                results = Calc.find_all(srch=srch, sd=sd)
                assert len(results) == 2
        """
        con = srch.findAll(sd)
        if con is None:
            print("Match result is null")
            return None
        c_count = con.getCount()
        if c_count == 0:
            print("No matches found")
            return None

        crs = []
        for i in range(c_count):
            try:
                cr = mLo.Lo.qi(XCellRange, con.getByIndex(i))
                if cr is None:
                    continue
                crs.append(cr)
            except Exception:
                print(f"Could not access match index {i}")
        if len(crs) == 0:
            print(f"Found {c_count} matches but unable to access any match")
            return None
        return crs

    # endregion ------------ search ------------------------------------

    # region --------------- cell decoration ---------------------------

    @staticmethod
    def create_cell_style(doc: XSpreadsheetDocument, style_name: str) -> XStyle | None:
        """
        Creates a style

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document
            style_name (str): Style name

        Returns:
            XStyle | None: Style on success; Otherwise, None
        """
        comp_doc = mLo.Lo.qi(XComponent, doc)
        style_families = mInfo.Info.get_style_container(doc=comp_doc, family_style_name="CellStyles")
        style = mLo.Lo.create_instance_msf(XStyle, "com.sun.star.style.CellStyle")
        #   "com.sun.star.sheet.TableCellStyle"  result in style == None ??
        try:
            style_families.insertByName(style_name, style)
            return style
        except Exception:
            print(f"Unable to create style: {style_name}")
        return None

    # region    change_style()

    @overload
    @staticmethod
    def change_style(sheet: XSpreadsheet, style_name: str, cell_range: XCellRange) -> bool:
        """
        Changes style fo a range of cells

        Args:
            sheet (XSpreadsheet): Spreadsheet
            style_name (str): Name of style to apply
            cell_range (XCellRange): Cell range to apply style to

        Returns:
            bool: True if style has been changed; Otherwise, False
        """
        ...

    @overload
    @staticmethod
    def change_style(sheet: XSpreadsheet, style_name: str, range_name: str) -> bool:
        """
        Changes style for a range of cells

        Args:
            sheet (XSpreadsheet): Spreadsheet
            style_name (str) :Name of style to apply
            range_name (str): Range to apply style to such as 'A1:E23'

        Returns:
            bool: True if style has been changed; Otherwise, False
        """
        ...

    @overload
    @staticmethod
    def change_style(
        sheet: XSpreadsheet, style_name: str, start_col: int, start_row: int, end_col: int, end_row: int
    ) -> bool:
        """
        Changes style fo a range of cells

        Args:
            sheet (XSpreadsheet): Spreadsheet
            style_name (str):  Name of style to apply
            start_col (int): Zero-base start column index
            start_row (int): Zero-base start row index
            end_col (int): Zero-base end column index
            end_row (int): Zero-base end row index

        Returns:
            bool: True if style has been changed; Otherwise, False
        """
        ...

    @classmethod
    def change_style(cls, *args, **kwargs) -> bool:
        ordered_keys = (1, 2, 3, 4, 5, 6)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('sheet', 'style_name', 'range_name', 'cell_range', 'start_col', 'start_row', 'end_col', 'end_row')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("change_style() got an unexpected keyword argument")
            ka[1] = kwargs.get("sheet", None)
            ka[2] = kwargs.get("style_name", None)
            keys = ("range_name", "start_col", "cell_range")
            for key in keys:
                if key in kwargs:
                    ka[3] = kwargs[key]
                    break
            if count == 3:
                return ka
            ka[4] = kwargs.get("start_row", None)
            ka[5] = kwargs.get("end_col", None)
            ka[6] = kwargs.get("end_row", None)
            return ka

        if not count in (3, 6):
            raise TypeError("change_style() got an invalid numer of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        if count == 3:
            if isinstance(kargs[3], str):
                # change_style(sheet: XSpreadsheet, style_name: str, range_name: str)
                cell_range = cls._get_cell_range_rng_name(sheet=kargs[1], range_name=kargs[3])  # 1 sheet, 3 range_name
                if cell_range is None:
                    return False
            else:
                cell_range = kargs[3]
            mProps.Props.set_property(prop_set=cell_range, name="CellStyle", value=kargs[2])  # 2 style_name
            return kargs[2] == mProps.Props.get_property(xprops=cell_range, name="CellStyle")
        else:
            # def change_style(sheet: XSpreadsheet, style_name: str, x1: int, y1: int, x2: int, y2:int)
            cell_range = cls._get_cell_range_col_row(
                sheet=kargs[1], start_col=kargs[3], start_row=kargs[4], end_col=kargs[5], end_row=kargs[6]
            )
            mProps.Props.set_property(prop_set=cell_range, name="CellStyle", value=kargs[2])  # 2 style_name
            return kargs[2] == mProps.Props.get_property(xprops=cell_range, name="CellStyle")

        # endregion change_style()

    # region    add_border()
    @classmethod
    def _add_border_sht_rng(cls, cell_range: XCellRange) -> None:
        cls._add_border_sht_rng_color(cell_range=cell_range, color=CommonColor.BLACK)  # color black

    @classmethod
    def _add_border_sht_rng_color(cls, cell_range: XCellRange, color: int) -> None:
        vals = (
            cls.BorderEnum.LEFT_BORDER
            | cls.BorderEnum.RIGHT_BORDER
            | cls.BorderEnum.TOP_BORDER
            | cls.BorderEnum.BOTTOM_BORDER
        )
        cls._add_border_sht_rng_color_vals(cell_range=cell_range, color=color, border_vals=vals)

    @classmethod
    def _add_border_sht_rng_color_vals(
        cls,
        cell_range: XCellRange,
        color: int,
        border_vals: int | BorderEnum,
    ) -> None:
        line = BorderLine2()  # create the border line
        line.Color = color
        line.InnerLineWidth = 0
        line.LineDistance = 0
        line.OuterLineWidth = 100
        bvs = cls.BorderEnum(int(border_vals))
        border = TableBorder2()

        if (bvs & cls.BorderEnum.TOP_BORDER) == cls.BorderEnum.TOP_BORDER:
            border.TopLine = line
            border.IsTopLineValid = True

        if (bvs & cls.BorderEnum.BOTTOM_BORDER) == cls.BorderEnum.BOTTOM_BORDER:
            border.BottomLine = line
            border.IsBottomLineValid = True

        if (bvs & cls.BorderEnum.LEFT_BORDER) == cls.BorderEnum.LEFT_BORDER:
            border.LeftLine = line
            border.IsLeftLineValid = True

        if (bvs & cls.BorderEnum.RIGHT_BORDER) == cls.BorderEnum.RIGHT_BORDER:
            border.RightLine = line
            border.IsRightLineValid = True
        mProps.Props.set_property(prop_set=cell_range, name="TableBorder2", value=border)

    @overload
    @staticmethod
    def add_border(sheet: XSpreadsheet, cell_range: XCellRange) -> XCellRange | None:
        """
        Adds borders to cell range

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cell_range (XCellRange):  Cell range

        Returns:
            XCellRange | None: Range that borders were applied to on success; Otherwise, None
        """
        ...

    @overload
    @staticmethod
    def add_border(sheet: XSpreadsheet, range_name: str) -> XCellRange | None:
        """
        Adds borders to cell range

        Args:
            sheet (XSpreadsheet): Spreadsheet
            range_name (str): Range Name such as 'A1:F9'

        Returns:
            XCellRange | None: Range that borders were applied to on success; Otherwise, None
        """
        ...

    @overload
    @staticmethod
    def add_border(sheet: XSpreadsheet, cell_range: XCellRange, color: int) -> XCellRange | None:
        """
        Adds borders to cell range

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cell_range (XCellRange): Cell range
            color (int): RGB color as integer

        Returns:
            XCellRange | None: Range that borders were applied to on success; Otherwise, None
        """
        ...

    @overload
    @staticmethod
    def add_border(sheet: XSpreadsheet, range_name: str, color: int) -> XCellRange | None:
        """
        Adds borders to cell range

        Args:
            sheet (XSpreadsheet): Spreadsheet
            range_name (str): Range Name such as 'A1:F9'
            color (int): RGB color as integer

        Returns:
            XCellRange | None: Range that borders were applied to on success; Otherwise, None
        """
        ...

    @overload
    @staticmethod
    def add_border(sheet: XSpreadsheet, cell_range: XCellRange, color: int, border_vals: int) -> XCellRange | None:
        """
        Adds borders to cell range

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cell_range (XCellRange): Cell range
            color (int): RGB color as integer
            border_vals (int): Determines what borders are applied.

        Returns:
            XCellRange | None: Range that borders were applied to on success; Otherwise, None
        """
        ...

    @overload
    @staticmethod
    def add_border(sheet: XSpreadsheet, range_name: str, color: int, border_vals: int) -> XCellRange | None:
        """
        Adds borders to cell range

        Args:
            sheet (XSpreadsheet): Spreadsheet
            range_name (str): Range Name such as 'A1:F9'
            color (int):  RGB color as integer
            border_vals (int): Determines what borders are applied.

        Returns:
            XCellRange | None: Range that borders were applied to on success; Otherwise, None
        """
        ...

    @overload
    @staticmethod
    def add_border(sheet: XSpreadsheet, cell_range: XCellRange, color: int, border_vals: BorderEnum) -> XCellRange | None:
        """
        Adds borders to cell range

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cell_range (XCellRange): Cell range
            color (int): RGB color as integer
            border_vals (BorderEnum): Determines what borders are applied.

        Returns:
            XCellRange | None: Range that borders were applied to on success; Otherwise, None
        """
        ...

    @overload
    @staticmethod
    def add_border(sheet: XSpreadsheet, range_name: str, color: int, border_vals: BorderEnum) -> XCellRange | None:
        """
        Adds borders to cell range

        Args:
            sheet (XSpreadsheet): Spreadsheet
            range_name (str): Range Name such as 'A1:F9'
            color (int): RGB color as integer
            border_vals (BorderEnum): Determines what borders are applied.

        Returns:
            XCellRange | None: Range that borders were applied to on success; Otherwise, None
        """
        ...

    @classmethod
    def add_border(cls, *args, **kwargs) -> XCellRange | None:
        ordered_keys = (1, 2, 3, 4)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('sheet', 'range_name', 'cell_range', 'color', 'border_vals')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("add_border() got an unexpected keyword argument")
            ka[1] = kwargs.get("sheet", None)
            keys = ("range_name", "cell_range")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("color", None)
            if count == 3:
                return ka
            ka[4] = kwargs.get("border_vals", None)
            return ka

        if not count in (2, 3, 4):
            raise TypeError("add_border() got an invalid numer of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        sheet = cast(XSpreadsheet, kargs[1])
        if isinstance(kargs[2], str):
            cell_range = sheet.getCellRangeByName(kargs[2])
        else:
            cell_range = kargs[2]

        if count == 2:
            # add_border(sheet: XSpreadsheet, cell_range: str)
            # add_border(sheet: XSpreadsheet, range_name: str)
            cls._add_border_sht_rng(cell_range=cell_range)
        elif count == 3:
            # add_border(sheet: XSpreadsheet, cell_range: XCellRange, color: int)
            #  add_border(sheet: XSpreadsheet, range_name: str, color: int)
            cls._add_border_sht_rng_color(cell_range=cell_range, color=kargs[3])
        else:
            # add_border(sheet: XSpreadsheet, cell_range: XCellRange, color: int, border_vals: int)
            # add_border(sheet: XSpreadsheet, range_name: str, color: int, border_vals: int)
            # add_border(sheet: XSpreadsheet, cell_range: XCellRange, color: int, border_vals: Calc.BorderEnum)
            # add_border(sheet: XSpreadsheet, range_name: str, color: int, border_vals: BorderEnum)
            cls._add_border_sht_rng_color_vals(cell_range=cell_range, color=kargs[3], border_vals=kargs[4])
        return cell_range

    # endregion add_border()
    @overload
    @staticmethod
    def highlight_range(sheet: XSpreadsheet,  headline: str, cell_range: XCellRange) ->  XCell | None:
        """
        Draw a colored border around the range and write a headline in the
        top-left cell of the range.

        Args:
            sheet (XSpreadsheet): Spreadsheet
            headline (str): Headline
            cell_range (XCellRange): Cell Range

        Returns:
            XCell | None: First cell of range that headline was applied to on success; Otherwise, None
        """
        ...
    
    @overload
    @staticmethod
    def highlight_range(sheet: XSpreadsheet,  headline: str, range_name: str) ->  XCell | None:
        """
        Draw a colored border around the range and write a headline in the
        top-left cell of the range.

        Args:
            sheet (XSpreadsheet): Spreadsheet
            headline (str): Headline
            range_name (str): Range Name such as 'A1:F9'

        Returns:
            XCell | None: First cell of range that headline was applied to on success; Otherwise, None            
        """
        ...

    @classmethod
    def highlight_range(cls, *args, **kwargs) -> XCell | None:
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('sheet', 'headline', 'range_name', 'cell_range')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("highlight_range() got an unexpected keyword argument")
            ka[1] = kwargs.get("sheet", None)
            ka[2] = kwargs.get("headline", None)
            keys = ("range_name", "cell_range")
            for key in keys:
                if key in kwargs:
                    ka[3] = kwargs[key]
                    break
            return ka

        if count != 3:
            raise TypeError("highlight_range() got an invalid numer of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        sheet = cast(XSpreadsheet, kargs[1])
        if isinstance(kargs[3], str):
            cell_range = sheet.getCellRangeByName(kargs[3])
        else:
            cell_range = kargs[3]

        cls._add_border_sht_rng_color(cell_range=cell_range, color=CommonColor.LIGHT_BLUE)

        # color the headline
        addr = cls._get_address_cell(cell_range=cell_range)
        if addr is None:
            return None
        header_range = cls._get_cell_range_col_row(
            sheet=sheet,
            start_col=addr.StartColumn,
            start_row=addr.StartRow,
            end_col=addr.EndColumn,
            end_row=addr.StartRow,
        )
        if header_range is None:
            return None
        first_cell = cls._get_cell_cell_rng(cell_range=header_range, column=0, row=0)
        if first_cell is None:
            return None
        cls._set_val_by_cell(value=kargs[2], cell=first_cell)
        return first_cell

    @classmethod
    def set_col_width(cls, sheet: XSpreadsheet, width: int, idx: int) -> XCellRange | None:
        """
        Sets column width. width is in ``mm``, e.g. 6

        Args:
            sheet (XSpreadsheet): Spreadsheet
            width (int): Width in mm
            idx (int): Index of column

        Returns:
            XCellRange | None: Column cell range that width is applied to on success; Otherwise, None
        """
        if width <= 0:
            print("Width must be greater then 0")
            return None
        cell_range = cls.get_col_range(sheet=sheet, idx=idx)
        if cell_range is None:
            return None
        mProps.Props.set_property(prop_set=cell_range, name="Width", value=(width * 100))
        return cell_range

    @classmethod
    def set_row_height(cls, sheet: XSpreadsheet, height: int, idx: int,) ->  XCellRange | None:
        """
        Sets column width. height is in ``mm``, e.g. 6

        Args:
            sheet (XSpreadsheet): Spreadsheet
            height (int): Width in mm
            idx (int): Index of Row

        Returns:
            XCellRange | None: Row cell range that height is applied to on success; Otherwise, None
        """
        if height <= 0:
            print("Height must be greater then 0")
            return None
        cell_range = cls.get_row_range(sheet=sheet, idx=idx)
        if cell_range is None:
            return
        # mInfo.Info.show_services(obj_name="Cell range for a row", obj=cell_range)
        mProps.Props.set_property(prop_set=cell_range, name="Height", value=(height * 100))
        return cell_range

    # endregion ------------ cell decoration ---------------------------

    # region --------------- scenarios ---------------------------------

    @staticmethod
    def insert_scenario(
        sheet: XSpreadsheet, range_str: str, vals: Sequence[Sequence[object]], name: str, comment: str
    ) -> XScenario |None:
        # get the cell range with the given address
        cell_range = sheet.getCellRangeByName(range_str)

        # create the range address sequence
        addr = mLo.Lo.qi(XCellRangeAddressable, cell_range)
        if addr is None:
            return None
        cr_addr = [addr.getRangeAddress()]

        # create the scenario
        supp = mLo.Lo.qi(XScenariosSupplier, sheet)
        scens = supp.getScenarios()
        scens.addNewByName(name, cr_addr, comment)

        # insert the values into the range
        cr_data = mLo.Lo.qi(XCellRangeData, cell_range)
        cr_data.setDataArray(vals)
        
        supp = mLo.Lo.qi(XScenariosSupplier, sheet)
        scenarios = supp.getScenarios()
        return mLo.Lo.qi(XScenario, scenarios.getByName(name))

    @staticmethod
    def apply_scenario(sheet: XSpreadsheet, name: str) -> XScenario | None:
        try:
            # get the scenario set
            supp = mLo.Lo.qi(XScenariosSupplier, sheet)
            scenarios = supp.getScenarios()

            # get the scenario and activate it
            scenario = mLo.Lo.qi(XScenario, scenarios.getByName(name))

            scenario.apply()
            return scenario
        except Exception as e:
            print("Scenario could not be applied:")
            print(f"    {e}")

    # endregion ------------ scenarios ---------------------------------

    # region --------------- data pilot methods ------------------------

    @staticmethod
    def get_pilot_tables(sheet: XSpreadsheet) -> XDataPilotTables | None:
        db_supp = mLo.Lo.qi(XDataPilotTablesSupplier, sheet)
        if db_supp is None:
            print("No data pilot supplier found")
            return None
        dp_tables = db_supp.getDataPilotTables()
        if dp_tables is None:
            print("No data pilot tables found")
            return None
        return dp_tables

    @staticmethod
    def get_pilot_table(dp_tables: XDataPilotTables, name: str) -> XDataPilotTable | None:
        try:
            otable = dp_tables.getByName(name)
            if otable is None:
                print(f"Did not find data pilot table '{name}'")
                return None
            return mLo.Lo.qi(XDataPilotTable, otable)
        except Exception as e:
            print(f"Pilot table lookup failed for '{name}'")
            print(f"    {e}")

        return None

    # endregion ------------ data pilot methods ------------------------

    # region --------------- using calc functions ----------------------

    @classmethod
    def compute_function(cls, fn: GeneralFunction | str, cell_range: XCellRange) -> float:
        """
        Compuutes a Calc Function

        Args:
            fn (GeneralFunction | str): Function to calculate, GeneralFunction Enum value or String such as 'SUM' or 'MAX'
            cell_range (XCellRange): Cell range to apply function on.

        Returns:
            float: result of function if successful. If there is an errro then 0.0 is returned.
        """
        try:
            sheet_op = mLo.Lo.qi(XSheetOperation, cell_range)
            func = cls.GeneralFunction(fn)  # convert to enum value if str
            if not mInfo.Info.is_type_enum(func, cls.GeneralFunction.__typename__):
                print("Arg fn is invalid, returning 0.0")
                return 0.0
            return sheet_op.computeFunction(func)
        except Exception as e:
            print("Compute function failed. Returning 0.0")
            print(f"    {e}")
        return 0.0

    # region    call_fun()

    @overload
    @staticmethod
    def call_fun(func_name: str, arg: object) -> object:
        ...

    @overload
    @staticmethod
    def call_fun(func_name: str, args: Tuple[object, ...]) -> object:
        ...

    @staticmethod
    def call_fun(*args, **kwargs) -> object | None:
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('func_name', 'arg', 'args')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("call_fun() got an unexpected keyword argument")
            ka[1] = kwargs.get("func_name", None)
            keys = ("arg", "args")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            return ka

        if count != 2:
            raise TypeError("call_fun() got an invalid numer of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if isinstance(kargs[2], tuple):
            arg = kargs[2]
        else:
            arg = [kargs[2]]
        try:
            fa = mLo.Lo.create_instance_mcf(XFunctionAccess, "com.sun.star.sheet.FunctionAccess")
            return fa.callFunction(kargs[1], kargs[2])
        except Exception:
            print(f"Could not invoke function '{kargs[1]}'")
        return None

    # endregion call_fun()

    @staticmethod
    def get_function_names() -> List[str] | None:
        funcs_desc = mLo.Lo.create_instance_mcf(XFunctionDescriptions, "com.sun.star.sheet.FunctionDescriptions")
        if funcs_desc is None:
            print("No function descriptions were found")
            return None

        nms: List[str] = []
        for i in range(funcs_desc.getCount()):
            try:
                props = cast(Sequence[PropertyValue], funcs_desc.getByIndex(i))
                for p in props:
                    if p.Name == "Name":
                        nms.append(str(p.Value))
                        break
            except Exception:
                continue
        if len(nms) == 0:
            print("No function names were found")
            return None
        nms.sort()
        return nms

    # region    find_function()

    @staticmethod
    def _find_function_by_name(func_nm: str) -> Tuple[PropertyValue] | None:
        if not func_nm:
            print("Please supply a function name to find.")
            return None
        func_desc = mLo.Lo.create_instance_mcf(XFunctionDescriptions, "com.sun.star.sheet.FunctionDescriptions")
        if func_desc is None:
            print("No function descriptions were found")
            return None

        for i in range(func_desc.getCount()):
            try:
                props = cast(Sequence[PropertyValue], func_desc.getByIndex(i))
                for p in props:
                    if p.Name == "Name" and str(p.Value) == func_nm:
                        return tuple(props)
            except Exception:
                continue
        print(f"Function '{func_nm}' not found")
        return None

    @staticmethod
    def _find_function_by_idx(idx: int) -> Tuple[PropertyValue] | None:
        if idx < 0:
            print("Please supply a positive index value to.")
            return None
        func_desc = mLo.Lo.create_instance_mcf(XFunctionDescriptions, "com.sun.star.sheet.FunctionDescriptions")
        if func_desc is None:
            print("No function descriptions were found")
            return None

        try:
            return tuple(func_desc.getByIndex(idx))
        except Exception as e:
            print(f"Could not access function description {idx}")
            print(f"    {e}")
        return None

    @overload
    @staticmethod
    def find_function(func_nm: str) -> Tuple[PropertyValue] | None:
        ...

    @overload
    @staticmethod
    def find_function(idx: int) -> Tuple[PropertyValue] | None:
        ...

    @classmethod
    def find_function(cls, *args, **kwargs) -> Tuple[PropertyValue] | None:
        ordered_keys = (1,)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('func_nm', 'idx')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("find_function() got an unexpected keyword argument")
            keys = ("func_nm", "idx")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            return ka

        if count != 1:
            raise TypeError("find_function() got an invalid numer of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        if isinstance(kargs[1], int):
            return cls._find_function_by_idx(kargs[1])
        return cls._find_function_by_name(kargs[1])

    # endregion find_function()

    @classmethod
    def print_function_info(cls, func_name: str) -> None:
        prop_vals = cls._find_function_by_name(func_nm=func_name)
        if prop_vals is None:
            return
        mProps.Props.show_props(prop_kind=func_name, props_set=prop_vals)
        cls.print_fun_arguments(prop_vals)
        print()

    @classmethod
    def print_fun_arguments(cls, prop_vals: Sequence[PropertyValue]) -> None:
        fargs: Sequence[FunctionArgument] = mProps.Props.get_value(name="Arguments", props=prop_vals)
        if fargs is None:
            print("No arguments found")
            return

        print(f"No. of arguments: {len(fargs)}")
        for i, arg in enumerate(fargs):
            cls.print_fun_argument(i, arg)

    @staticmethod
    def print_fun_argument(i: int, fa: FunctionArgument) -> None:
        print(f"{i+1}. Argument name: {fa.Name}")
        print(f"  Description: '{fa.Description}'")
        print(f"  Is optional?: {fa.IsOptional}")
        print()

    @staticmethod
    def get_recent_functions() -> Tuple[int, ...] | None:
        recent_funcs = mLo.Lo.create_instance_mcf(XRecentFunctions, "com.sun.star.sheet.RecentFunctions")
        if recent_funcs is None:
            print("No recent functions found")
            return None

        return recent_funcs.getRecentFunctionIds()

    # endregion ------------ using calc functions ----------------------

    # region --------------- solver methods ----------------------------

    @classmethod
    def goal_seek(
        cls, gs: XGoalSeek, sheet: XSpreadsheet, cell_name: str, formula_cell_name: str, result: float
    ) -> float:
        """find x in formula when it equals result"""
        xpos = cls._get_cell_address_sheet(sheet=sheet, cell_name=cell_name)
        formula_pos = cls._get_cell_address_sheet(sheet=sheet, cell_name=formula_cell_name)

        goal_result = gs.seekGoal(formula_pos, xpos, f"{result}")
        if goal_result.Divergence >= 0.1:
            print(f"NO result; divergence: {goal_result.Divergence}")
        return goal_result.Divergence

    @staticmethod
    def list_solvers() -> None:
        print("Services offered by the solver:")
        nms = mInfo.Info.get_service_names(service_name="com.sun.star.sheet.Solver")
        if nms is None:
            print("  none")
            return

        for service in nms:
            print(f"  {service}")
        print()

    @classmethod
    def to_constraint_op(cls, op: str) -> SolverConstraintOperator:
        if op == "=" or op == "==":
            return cls.SolverConstraintOperator.EQUAL
        if (op == "<=") or op == "=<":
            return cls.SolverConstraintOperator.LESS_EQUAL
        if (op == ">=") or op == "=>":
            return cls.SolverConstraintOperator.GREATER_EQUAL
        print(f"Do not recognise op: {op}; using EQUAL")
        return cls.SolverConstraintOperator.EQUAL

    # region    make_constraint()
    @classmethod
    def _make_constraint_op_str_sht_cell_name(
        cls, num: float, op: str, sheet: XSpreadsheet, cell_name: str
    ) -> SolverConstraint | None:
        return cls._make_constraint_op_str_addr(
            num=num, op=op, addr=cls._get_cell_address_sheet(sheet=sheet, cell_name=cell_name)
        )

    @classmethod
    def _make_constraint_op_str_addr(cls, num: float, op: str, addr: CellAddress) -> SolverConstraint | None:
        return cls._make_constraint_op_sco_addr(num=num, op=cls.to_constraint_op(op), addr=addr)

    @classmethod
    def _make_constraint_op_sco_sht_cell_name(
        cls, num: float, op: SolverConstraintOperator, sheet: XSpreadsheet, cell_name: str
    ) -> SolverConstraint | None:
        return cls._make_constraint_op_sco_addr(
            num=num, op=op, addr=cls._get_cell_address_sheet(sheet=sheet, cell_name=cell_name)
        )

    @classmethod
    def _make_constraint_op_sco_addr(
        cls, num: float, op: SolverConstraintOperator, addr: CellAddress
    ) -> SolverConstraint | None:
        sc = SolverConstraint()
        sc.Left = addr
        sc.Operator = op
        sc.Right = num
        return sc

    @overload
    @staticmethod
    def make_constraint(num: float, op: str, addr: CellAddress) -> SolverConstraint | None:
        ...

    @overload
    @staticmethod
    def make_constraint(num: float, op: SolverConstraintOperator, addr: CellAddress) -> SolverConstraint | None:
        ...

    @overload
    @staticmethod
    def make_constraint(num: float, op: str, sheet: XSpreadsheet, cell_name: str) -> SolverConstraint | None:
        ...

    @overload
    @staticmethod
    def make_constraint(
        num: float, op: SolverConstraintOperator, sheet: XSpreadsheet, cell_name: str
    ) -> SolverConstraint | None:
        ...

    @classmethod
    def make_constraint(cls, *args, **kwargs) -> SolverConstraint | None:
        ordered_keys = (1, 2, 3, 4)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ('num', 'op', 'sheet', 'addr', 'cell_name')
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("make_constraint() got an unexpected keyword argument")
            ka[1] = kwargs.get("num", None)
            ka[2] = kwargs.get("op", None)
            keys = ("sheet", "addr")
            for key in keys:
                if key in kwargs:
                    ka[3] = kwargs[key]
                    break
            if count == 3:
                return ka
            ka[4] = kwargs.get("cell_name", None)
            return ka

        if not count(3, 4):
            raise TypeError("make_constraint() got an invalid numer of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        if count == 3:
            if isinstance(kargs[2], str):
                # def make_constraint(num: float, op: str, addr: CellAddress)
                return cls._make_constraint_op_str_addr(num=kargs[1], op=kargs[2], addr=kargs[3])
            else:
                # def make_constraint(num: float, op: SolverConstraintOperator, addr: CellAddress)
                return cls._make_constraint_op_sco_addr(num=kargs[1], op=kargs[2], addr=kargs[3])
        else:
            if isinstance(kargs[2], str):
                # def make_constraint(num: float, op: str, sheet: XSpreadsheet, cell_name:str)
                return cls._make_constraint_op_str_sht_cell_name(
                    num=kargs[1], op=kargs[2], sheet=kargs[3], cell_name=kargs[4]
                )
            else:
                # def make_constraint(num: float, op: SolverConstraintOperator, sheet: XSpreadsheet, cell_name:str)
                return cls._make_constraint_op_sco_sht_cell_name(
                    num=kargs[1], op=kargs[2], sheet=kargs[3], cell_name=kargs[4]
                )

    # endregion    make_constraint()

    @classmethod
    def solver_report(cls, solver: XSolver) -> None:
        # note: in original java it was getSuccess(), getObjective(), getVariables(), getSolution(),
        # These are typedef properties. The types-unopy typings are correct. Typedef are represented as Class Properties.
        is_successful = solver.Success
        cell_name = cls._get_cell_str_addr(solver.Objective)
        print("Solver result: ")
        print(f"  {cell_name} == {solver.ResultValue:.4f}")
        addrs = solver.Variables
        solns = solver.Solution
        print("Solver variables: ")
        for i, num in enumerate(solns):
            cell_name = cls._get_cell_str_addr(addrs[i])
            print(f"  {cell_name} == {num:.4f}")
        print()

    # endregion ------------ solver methods ----------------------------

    # region --------------- headers /footers --------------------------

    @staticmethod
    def get_head_foot(props: XPropertySet, content: str) -> XHeaderFooterContent | None:
        return mLo.Lo.qi(XHeaderFooterContent, mProps.Props.get_property(xprops=props, name=content))

    @staticmethod
    def print_head_foot(title: str, hfc: XHeaderFooterContent) -> None:
        left = hfc.getLeftText()
        center = hfc.getCenterText()
        right = hfc.getRightText()
        print(f"{title}: '{left.getString()}' : '{center.getString()}' : '{right.getString()}'")

    @classmethod
    def get_region(cls, hfc: XHeaderFooterContent, region: int) -> XText | None:
        if hfc is None:
            print("Header/footer content is null")
            return None

        if region == cls.HeaderFooter.HF_LEFT:
            return hfc.getLeftText()
        if region == cls.HeaderFooter.HF_CENTER:
            return hfc.getCenterText()
        if region == cls.HeaderFooter.HF_RIGHT:
            return hfc.getRightText()
        print("Unknown header/footer region")
        return None

    @classmethod
    def set_head_foot(cls, hfc: XHeaderFooterContent, region: int, text: str) -> None:
        xtext = cls.get_region(hfc=hfc, region=region)
        if xtext is None:
            print("Could not set text")
            return
        header_cursor = xtext.createTextCursor()
        header_cursor.gotoStart(False)
        header_cursor.gotoEnd(True)
        header_cursor.setString(text)

    # endregion --------------- headers /footers -----------------------
