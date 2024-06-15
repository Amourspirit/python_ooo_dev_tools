# coding: utf-8
# Python conversion of Calc.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/
# region Imports
from __future__ import annotations
import contextlib
import itertools
from enum import IntEnum, IntFlag, Enum
import re
from typing import Any, List, Tuple, cast, overload, Sequence, Optional, TYPE_CHECKING
import uno

# from ..mock import mock_g

# if not mock_g.DOCS_BUILDING:
# not importing for doc building just result in short import name for
# args that use these.
# this is also true because docs/conf.py ignores com import for autodoc
from com.sun.star.beans import XPropertySet
from com.sun.star.container import XEnumerationAccess
from com.sun.star.container import XIndexAccess
from com.sun.star.container import XNamed
from com.sun.star.frame import XModel
from com.sun.star.lang import Locale
from com.sun.star.lang import XComponent
from com.sun.star.sheet import SolverConstraint  # struct
from com.sun.star.sheet import XCellAddressable
from com.sun.star.sheet import XCellRangeAddressable
from com.sun.star.sheet import XCellRangeData
from com.sun.star.sheet import XCellRangeMovement
from com.sun.star.sheet import XCellRangesQuery
from com.sun.star.sheet import XCellSeries
from com.sun.star.sheet import XDataPilotTable
from com.sun.star.sheet import XDataPilotTablesSupplier
from com.sun.star.sheet import XFunctionAccess
from com.sun.star.sheet import XFunctionDescriptions
from com.sun.star.sheet import XHeaderFooterContent
from com.sun.star.sheet import XPrintAreas
from com.sun.star.sheet import XRecentFunctions
from com.sun.star.sheet import XScenario
from com.sun.star.sheet import XScenariosSupplier
from com.sun.star.sheet import XSheetAnnotationAnchor
from com.sun.star.sheet import XSheetAnnotationsSupplier
from com.sun.star.sheet import XSheetCellRange
from com.sun.star.sheet import XSheetOperation
from com.sun.star.sheet import XSpreadsheet
from com.sun.star.sheet import XSpreadsheetDocument
from com.sun.star.sheet import XSpreadsheets
from com.sun.star.sheet import XSpreadsheetView
from com.sun.star.sheet import XUsedAreaCursor
from com.sun.star.sheet import XViewFreezable
from com.sun.star.sheet import XViewPane
from com.sun.star.style import XStyle
from com.sun.star.table import BorderLine2  # struct
from com.sun.star.table import TableBorder2  # struct
from com.sun.star.table import XCell
from com.sun.star.table import XCellRange
from com.sun.star.table import XColumnRowRange
from com.sun.star.text import XSimpleText
from com.sun.star.uno import Exception as UnoException
from com.sun.star.util import NumberFormat  # const
from com.sun.star.util import XMergeable
from com.sun.star.util import XNumberFormatsSupplier
from com.sun.star.util import XNumberFormatTypes
from com.sun.star.util import XProtectable
from com.sun.star.view import XPrintable


from ooo.dyn.awt.point import Point
from ooo.dyn.beans.property_value import PropertyValue
from ooo.dyn.lang.illegal_argument_exception import IllegalArgumentException
from ooo.dyn.sheet.cell_delete_mode import CellDeleteMode
from ooo.dyn.sheet.cell_flags import CellFlagsEnum as CellFlagsEnum
from ooo.dyn.sheet.cell_insert_mode import CellInsertMode
from ooo.dyn.sheet.fill_date_mode import FillDateMode as FillDateMode
from ooo.dyn.sheet.general_function import GeneralFunction as GeneralFunction
from ooo.dyn.sheet.solver_constraint_operator import SolverConstraintOperator as SolverConstraintOperator
from ooo.dyn.table.cell_content_type import CellContentType
from ooo.dyn.table.cell_hori_justify import CellHoriJustify
from ooo.dyn.table.cell_vert_justify2 import CellVertJustify2

from ooodev.exceptions import ex as mEx
from ooodev.formatters.formatter_table import FormatterTable
from ooodev.units.unit_convert import UnitConvert
from ooodev.gui import gui as mGui
from ooodev.utils import info as mInfo
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.utils import table_helper as mTblHelper
from ooodev.utils import view_state as mViewState
from ooodev.utils.color import CommonColor, Color
from ooodev.utils.data_type import cell_obj as mCellObj
from ooodev.utils.data_type import range_obj as mRngObj
from ooodev.utils.data_type import range_values as mRngValues
from ooodev.utils.data_type.size import Size
from ooodev.utils.gen_util import ArgsHelper, Util as GenUtil
from ooodev.utils.kind.zoom_kind import ZoomKind
from ooodev.utils.type_var import PathOrStr, Row, Column, Table, TupleArray, FloatList, FloatTable
from ooodev.events.args.calc.cell_args import CellArgs
from ooodev.events.args.calc.cell_cancel_args import CellCancelArgs
from ooodev.events.args.calc.sheet_args import SheetArgs
from ooodev.events.args.calc.sheet_cancel_args import SheetCancelArgs
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.calc_named_event import CalcNamedEvent
from ooodev.events.event_singleton import _Events

if TYPE_CHECKING:
    from com.sun.star.frame import XComponentLoader
    from com.sun.star.frame import XController
    from com.sun.star.frame import XFrame
    from com.sun.star.sheet import FunctionArgument  # struct
    from com.sun.star.sheet import XDataPilotTables
    from com.sun.star.sheet import XGoalSeek
    from com.sun.star.sheet import XSheetAnnotation
    from com.sun.star.sheet import XSheetCellCursor
    from com.sun.star.sheet import XSolver
    from com.sun.star.table import CellAddress

    # from com.sun.star.table import CellRangeAddress
    from ooo.dyn.table.cell_range_address import CellRangeAddress
    from com.sun.star.text import XText
    from com.sun.star.util import XSearchable
    from com.sun.star.util import XSearchDescriptor
    from com.sun.star.util import CellProtection
    from ooodev.units.unit_obj import UnitT
    from ooodev.proto.style_obj import StyleT
else:
    XComponentLoader = Any
    XController = Any
    XFrame = Any
    FunctionArgument = Any
    XDataPilotTables = Any
    XGoalSeek = Any
    XSheetAnnotation = Any
    XSheetCellCursor = Any
    XSolver = Any
    CellAddress = Any
    CellRangeAddress = Any
    XText = Any
    XSearchable = Any
    XSearchDescriptor = Any
    CellProtection = Any
    UnitT = Any
    StyleT = Any

NameVal = ArgsHelper.NameValue
# endregion Imports


class Calc:
    # region classes
    # for headers and footers
    class HeaderFooter(IntEnum):
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

        def __str__(self) -> str:
            return self.value

    # endregion classes

    # region Constants
    # largest value used in XCellSeries.fillSeries
    MAX_VALUE = 0x7FFFFFFF

    # use a better name when date mode doesn't matter
    NO_DATE = FillDateMode.FILL_DATE_DAY

    CELL_POS = Point(3, 4)

    _rx_cell = re.compile(r"([a-zA-Z]+)([0-9]+)")

    # endregion Constants

    # region --------------- document methods --------------------------

    # region open_doc()
    @overload
    @classmethod
    def open_doc(cls) -> XSpreadsheetDocument: ...

    @overload
    @classmethod
    def open_doc(cls, fnm: PathOrStr) -> XSpreadsheetDocument: ...

    @overload
    @classmethod
    def open_doc(cls, *, visible: bool) -> XSpreadsheetDocument: ...

    @overload
    @classmethod
    def open_doc(cls, fnm: PathOrStr, *, visible: bool) -> XSpreadsheetDocument: ...

    @overload
    @classmethod
    def open_doc(cls, *, loader: XComponentLoader) -> XSpreadsheetDocument: ...

    @overload
    @classmethod
    def open_doc(cls, *, loader: XComponentLoader, visible: bool) -> XSpreadsheetDocument: ...

    @overload
    @classmethod
    def open_doc(cls, fnm: PathOrStr, loader: XComponentLoader) -> XSpreadsheetDocument: ...

    @overload
    @classmethod
    def open_doc(cls, fnm: PathOrStr, loader: XComponentLoader, *, visible: bool) -> XSpreadsheetDocument: ...

    @classmethod
    def open_doc(
        cls, fnm: PathOrStr | None = None, loader: XComponentLoader | None = None, **kwargs: Any
    ) -> XSpreadsheetDocument:
        """
        Opens or creates a spreadsheet document.

        |lo_unsafe|

        Args:
            fnm (str): Spreadsheet file to open. If omitted then a new Spreadsheet document is returned.
            loader (XComponentLoader): Component loader

        Raises:
            CancelEventError: If ``DOC_OPENING`` is canceled

        Returns:
            XSpreadsheetDocument: Spreadsheet document

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.DOC_OPENING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.DOC_OPENED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing all method parameters.

            If ``fnm`` is omitted then ``DOC_OPENED`` event will not be raised.
        """
        # MacroExecutionMode=MacroExecMode.ALWAYS_EXECUTE
        props_dict = {"Hidden": True}
        if "visible" in kwargs:
            visible = bool(kwargs.pop("visible"))
            props_dict["Hidden"] = not visible
        if kwargs:
            props_dict.update(kwargs)
        local_props = mProps.Props.make_props(**props_dict)
        cargs = CancelEventArgs(Calc.open_doc.__qualname__)
        cargs.event_data = {"fnm": fnm, "loader": loader}
        _Events().trigger(CalcNamedEvent.DOC_OPENING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)
        if _fnm := cast(PathOrStr, cargs.event_data["fnm"]):
            doc = (
                mLo.Lo.open_doc(fnm=_fnm, props=local_props)
                if loader is None
                else mLo.Lo.open_doc(fnm=_fnm, loader=loader, props=local_props)
            )
            _Events().trigger(CalcNamedEvent.DOC_OPENED, EventArgs.from_args(cargs))
        elif loader is None:
            doc = cls.create_doc()
        else:
            doc = cls.create_doc(loader=loader)
        return cls.get_ss_doc(doc)  # type: ignore

    # endregion open_doc()

    @staticmethod
    def save_doc(doc: XSpreadsheetDocument, fnm: PathOrStr) -> bool:
        """
        Saves text document.

        |lo_safe|

        Args:
            text_doc (XSpreadsheetDocument): Text Document
            fnm (PathOrStr): Path to save as

        Raises:
            MissingInterfaceError: If doc does not implement XComponent interface

        Returns:
            bool: True if doc is saved; Otherwise, False

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.DOC_SAVING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.DOC_SAVED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing ``text_doc`` and ``fnm``.

        Attention:
            :py:meth:`Lo.save_doc <.utils.lo.Lo.save_doc>` method is called along with any of its events.
        """
        cargs = CancelEventArgs(Calc.save_doc.__qualname__)
        cargs.event_data = {"doc": doc, "fnm": fnm}
        _Events().trigger(CalcNamedEvent.DOC_SAVING, cargs)

        if cargs.cancel:
            return False
        fnm = cast(PathOrStr, cargs.event_data["fnm"])

        comp = mLo.Lo.qi(XComponent, doc, raise_err=True)
        result = mLo.Lo.save_doc(doc=comp, fnm=fnm)

        _Events().trigger(CalcNamedEvent.DOC_SAVED, EventArgs.from_args(cargs))
        return result

    @staticmethod
    def get_ss_doc(doc: XComponent) -> XSpreadsheetDocument:
        """
        Gets a spreadsheet document.

        When using this method in a macro the :py:attr:`Lo.this_component <.utils.lo.Lo.this_component>` value should be passed as ``doc`` arg.

        |lo_safe|

        Args:
            doc (XComponent): Component to get spreadsheet from

        Raises:
            Exception: If not a spreadsheet document
            MissingInterfaceError: If doc does not have XSpreadsheetDocument interface

        Returns:
            XSpreadsheetDocument: spreadsheet document

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.DOC_SS` :eventref:`src-docs-event`

        See Also:
            :py:meth:`~Calc.create_doc`
        """
        # sourcery skip: raise-specific-error
        if not mInfo.Info.is_doc_type(doc_type=mLo.Lo.Service.CALC, obj=doc):
            if not mLo.Lo.is_macro_mode:
                mLo.Lo.close_doc(doc=doc)
            raise Exception("Not a spreadsheet doc")

        ss_doc = mLo.Lo.qi(XSpreadsheetDocument, doc)
        if ss_doc is None:
            if not mLo.Lo.is_macro_mode:
                mLo.Lo.close_doc(doc=doc)
            raise mEx.MissingInterfaceError(XSpreadsheetDocument)
        _Events().trigger(CalcNamedEvent.DOC_SS, EventArgs(Calc.get_ss_doc.__qualname__))
        return ss_doc

    # region create_doc()
    @overload
    @staticmethod
    def create_doc() -> XSpreadsheetDocument: ...

    @overload
    @staticmethod
    def create_doc(loader: XComponentLoader) -> XSpreadsheetDocument: ...

    @overload
    @staticmethod
    def create_doc(*, visible: bool) -> XSpreadsheetDocument: ...

    @overload
    @staticmethod
    def create_doc(loader: XComponentLoader, *, visible: bool) -> XSpreadsheetDocument: ...

    @staticmethod
    def create_doc(loader: XComponentLoader | None = None, **kwargs: Any) -> XSpreadsheetDocument:
        """
        Creates a new spreadsheet document.

        |lo_unsafe|

        Args:
            loader (XComponentLoader): Component Loader.

        Raises:
            MissingInterfaceError: If doc does not have XSpreadsheetDocument interface.
            CancelEventError: If DOC_CREATING event is canceled.

        Returns:
            XSpreadsheetDocument: Spreadsheet document.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.DOC_CREATING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.DOC_CREATED` :eventref:`src-docs-event`

        See Also:
            :py:meth:`~Calc.get_ss_doc`

        Note:
            Event args ``event_data`` is a dictionary containing ``loader``.
        """

        # MacroExecutionMode=MacroExecMode.ALWAYS_EXECUTE

        props_dict = {"Hidden": True}
        if "visible" in kwargs:
            visible = bool(kwargs.pop("visible"))
            props_dict["Hidden"] = not visible
        if kwargs:
            props_dict.update(kwargs)
        local_props = mProps.Props.make_props(**props_dict)
        cargs = CancelEventArgs(Calc.create_doc.__qualname__)
        cargs.event_data = {"loader": loader}
        _Events().trigger(CalcNamedEvent.DOC_CREATING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)
        if loader is None:
            doc = mLo.Lo.qi(
                XSpreadsheetDocument, mLo.Lo.create_doc(doc_type=mLo.Lo.DocTypeStr.CALC, props=local_props), True
            )
        else:
            doc = mLo.Lo.qi(
                XSpreadsheetDocument,
                mLo.Lo.create_doc(doc_type=mLo.Lo.DocTypeStr.CALC, loader=loader, props=local_props),
                True,
            )
        _Events().trigger(CalcNamedEvent.DOC_CREATED, EventArgs.from_args(cargs))
        return doc

        # XSpreadsheetDocument does not inherit XComponent!

    # endregion create_doc()

    @classmethod
    def get_current_doc(cls) -> XSpreadsheetDocument:
        """
        Gets the current document.

        |lo_unsafe|

        Raises:
            NoneError: If no current document

        Returns:
            XSpreadsheetDocument: Spreadsheet Document
        """
        doc = mLo.Lo.this_component
        if doc is None:
            # most likely in headless mode with option dynamic set to True
            doc = mLo.Lo.lo_component
        if doc is None:
            raise mEx.NoneError("current document")
        return cls.get_ss_doc(doc)

    @classmethod
    def get_doc_from_sheet(cls, sheet: XSpreadsheetDocument) -> XSpreadsheetDocument:
        """
        Gets the document from a sheet.
        
        Args:
            sheet (XSpreadsheetDocument): Sheet to get document from.
        
        Returns:
            XSpreadsheetDocument: Spreadsheet Document.
        
        .. versionadded:: 0.46.0
        """
        sht = cast(Any, sheet)
        imp_name = sht.getImplementationName()
        if imp_name != "ScTableSheetObj":
            raise Exception("Not a spreadsheet sheet")
        return sht.DrawPage.Forms.Parent

    # endregion ------------ document methods ------------------

    # region --------------- sheet methods -----------------------------

    # region    get_sheet()
    @staticmethod
    def _get_sheet_index(doc: XSpreadsheetDocument, index: int) -> XSpreadsheet:
        """LO Safe Method. Return the spreadsheet with the specified index (0-based)"""
        # sourcery skip: raise-specific-error
        cargs = SheetCancelArgs(Calc.get_sheet.__qualname__)
        cargs.index = index
        cargs.name = None
        cargs.doc = doc

        _Events().trigger(CalcNamedEvent.SHEET_GETTING, cargs)
        if cargs.cancel:
            mEx.CancelEventError(cargs)

        index = cargs.index
        sheets = cargs.doc.getSheets()
        try:
            sheets_idx = mLo.Lo.qi(XIndexAccess, sheets, True)
            sheet = mLo.Lo.qi(XSpreadsheet, sheets_idx.getByIndex(index), raise_err=True)
            _Events().trigger(CalcNamedEvent.SHEET_GET, SheetArgs.from_args(cargs))
            return sheet
        except Exception as e:
            raise Exception(f"Could not access spreadsheet: {index}") from e

    @staticmethod
    def _get_sheet_name(doc: XSpreadsheetDocument, sheet_name: str) -> XSpreadsheet:
        """Lo Safe Method. Return the spreadsheet with the specified index (0-based)"""
        # sourcery skip: raise-specific-error
        cargs = SheetCancelArgs(Calc.get_sheet.__qualname__)
        cargs.name = sheet_name
        cargs.index = None
        cargs.doc = doc
        _Events().trigger(CalcNamedEvent.SHEET_GETTING, cargs)
        if cargs.cancel:
            mEx.CancelEventError(cargs)
        sheet_name = cargs.name
        sheets = cargs.doc.getSheets()
        try:
            sheet = mLo.Lo.qi(XSpreadsheet, sheets.getByName(sheet_name), raise_err=True)
            _Events().trigger(CalcNamedEvent.SHEET_GET, SheetArgs.from_args(cargs))
            return sheet
        except Exception as e:
            raise Exception(f"Could not access spreadsheet: '{sheet_name}'") from e

    @overload
    @classmethod
    def get_sheet(cls) -> XSpreadsheet: ...

    @overload
    @classmethod
    def get_sheet(cls, idx: int) -> XSpreadsheet: ...

    @overload
    @classmethod
    def get_sheet(cls, sheet_name: str) -> XSpreadsheet: ...

    @overload
    @classmethod
    def get_sheet(cls, doc: XSpreadsheetDocument) -> XSpreadsheet: ...

    @overload
    @classmethod
    def get_sheet(cls, doc: XSpreadsheetDocument, idx: int) -> XSpreadsheet: ...

    @overload
    @classmethod
    def get_sheet(cls, doc: XSpreadsheetDocument, sheet_name: str) -> XSpreadsheet: ...

    @classmethod
    def get_sheet(cls, *args, **kwargs) -> XSpreadsheet:
        """
        Gets a sheet of spreadsheet document.

        |lo_safe|

        Args:
            doc (XSpreadsheetDocument, optional): Spreadsheet document.
            idx (int, optional): Zero based index of spreadsheet. Defaults to ``0``.
            sheet_name (str, optional): Name of spreadsheet.

        Raises:
            Exception: If spreadsheet is not found.
            CancelEventError: If SHEET_GETTING event is canceled.

        Returns:
            XSpreadsheet: Spreadsheet at index.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_GETTING` :eventref:`src-docs-sheet-event-getting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_GET` :eventref:`src-docs-sheet-event-get`

        Note:
            For Event args, if ``index`` is available then ``name`` is ``None`` and if ``sheet_name`` is available then ``index`` is ``None``.

        .. versionchanged:: 0.6.10

            Added overload ``get_sheet(doc: XSpreadsheetDocument) -> XSpreadsheet``

        .. versionchanged:: 0.8.6
            Added overload ``get_sheet() -> XSpreadsheet``.
            Added overload ``get_sheet(idx: int) -> XSpreadsheet``.
            Added overload ``get_sheet(sheet_name: str) -> XSpreadsheet``.
            Changed ``get_sheet(doc: XSpreadsheetDocument, index: int)`` to ``get_sheet(doc: XSpreadsheetDocument, idx: int)``
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        # index is backwards compatibility
        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("doc", "idx", "index", "sheet_name")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_sheet() got an unexpected keyword argument")
            ka[1] = kwargs.get("doc", None)
            if count == 1:
                return ka
            keys = ("index", "idx", "sheet_name")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            return ka

        if count not in (0, 1, 2):
            raise TypeError("get_sheet() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 0:
            return cls._get_sheet_index(cls.get_current_doc(), 0)

        arg1 = kargs[1]
        if count == 1:
            if isinstance(arg1, int):
                return cls._get_sheet_index(cls.get_current_doc(), arg1)
            if isinstance(arg1, str):
                return cls._get_sheet_name(cls.get_current_doc(), arg1)
            return cls._get_sheet_index(arg1, 0)

        arg2 = kargs[2]
        if isinstance(arg2, int):
            return cls._get_sheet_index(arg1, arg2)

        return cls._get_sheet_name(arg1, arg2)

    # endregion get_sheet()

    @staticmethod
    def insert_sheet(doc: XSpreadsheetDocument, name: str, idx: int) -> XSpreadsheet:
        """
        Inserts a spreadsheet into document.

        |lo_safe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet document
            name (str): Name of sheet to insert
            idx (int): zero-based index position of the sheet to insert

        Raises:
            Exception: If unable to insert spreadsheet
            CancelEventError: If SHEET_INSERTING event is canceled

        Returns:
            XSpreadsheet: The newly inserted sheet

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_INSERTING` :eventref:`src-docs-sheet-event-inserting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_INSERTED` :eventref:`src-docs-sheet-event-inserted`
        """
        # sourcery skip: raise-specific-error
        cargs = SheetCancelArgs(Calc.insert_sheet.__qualname__)
        cargs.name = name
        cargs.index = idx
        cargs.doc = doc
        _Events().trigger(CalcNamedEvent.SHEET_INSERTING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)
        name = cargs.name
        idx = cargs.index
        sheets = cargs.doc.getSheets()
        try:
            sheets.insertNewByName(name, idx)
            sheet = mLo.Lo.qi(XSpreadsheet, sheets.getByName(name), raise_err=True)
            _Events().trigger(CalcNamedEvent.SHEET_INSERTED, SheetArgs.from_args(cargs))
            return sheet
        except Exception as e:
            raise Exception("Could not insert sheet:") from e

    # region    remove_sheet()

    @staticmethod
    def _remove_sheet_name(doc: XSpreadsheetDocument, sheet_name: str) -> bool:
        """LO Safe Method"""
        cargs = SheetCancelArgs(Calc.remove_sheet.__qualname__)
        # cargs.source = Calc.remove_sheet
        cargs.doc = doc
        cargs.name = sheet_name
        cargs.index = None
        cargs.event_data = {"fn_type": "name"}
        _Events().trigger(CalcNamedEvent.SHEET_REMOVING, cargs)
        if cargs.cancel:
            return False

        sheet_name = cargs.name
        sheets = cargs.doc.getSheets()
        result = False
        try:
            sheets.removeByName(sheet_name)
            result = True
        except Exception:
            mLo.Lo.print(f"Could not remove sheet: {sheet_name}")
        if result:
            _Events().trigger(CalcNamedEvent.SHEET_REMOVED, SheetArgs.from_args(cargs))
        return result

    @classmethod
    def _remove_sheet_index(cls, doc: XSpreadsheetDocument, index: int) -> bool:
        """Lo safe method"""
        cargs = SheetCancelArgs(Calc.remove_sheet.__qualname__)
        cargs.doc = doc
        cargs.index = index
        cargs.name = None
        cargs.event_data = {"fn_type": "index"}
        _Events().trigger(CalcNamedEvent.SHEET_REMOVING, cargs)
        if cargs.cancel:
            return False

        index = cargs.index
        sheets = cargs.doc.getSheets()
        result = False
        try:
            sheets_idx = mLo.Lo.qi(XIndexAccess, sheets, True)
            sheet = mLo.Lo.qi(XSpreadsheet, sheets_idx.getByIndex(index), True)
            sheet_name = cls.get_sheet_name(sheet, safe_quote=False)
            if sheet_name is None:
                return False
            sheets.removeByName(sheet_name)
            result = True
        except Exception:
            mLo.Lo.print(f"Could not remove sheet at index: {index}")
        if result:
            _Events().trigger(CalcNamedEvent.SHEET_REMOVED, SheetArgs.from_args(cargs))
        return result

    @overload
    @classmethod
    def remove_sheet(cls, doc: XSpreadsheetDocument, sheet_name: str) -> bool: ...

    @overload
    @classmethod
    def remove_sheet(cls, doc: XSpreadsheetDocument, idx: int) -> bool: ...

    @classmethod
    def remove_sheet(cls, *args, **kwargs) -> bool:
        """
        Removes a sheet from document.

        |lo_safe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet document.
            sheet_name (str): Name of sheet to remove.
            idx (int): Zero based index of sheet to remove.

        Returns:
            bool: ``True`` of sheet was removed; Otherwise, ``False``.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_REMOVING` :eventref:`src-docs-sheet-event-removing`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_REMOVED` :eventref:`src-docs-sheet-event-removed`

        Note:
            Event args ``event_data`` is set to a dictionary.
            If ``idx`` is available then args ``event_data["fn_type"]`` is set to a value ``idx``; Otherwise, set to a value ``name``.

        .. versionchanged:: 0.8.6
            Renamed ``index`` arg to ``idx``. ``index`` will still work but is now undocumented.
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        # index is backwards compatible
        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("doc", "index", "idx", "sheet_name")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("remove_sheet() got an unexpected keyword argument")
            ka[1] = kwargs.get("doc", None)
            keys = ("index", "idx", "sheet_name")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            return ka

        if count != 2:
            raise TypeError("remove_sheet() got an invalid number of arguments")

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
        Moves a sheet in a spreadsheet document.

        |lo_safe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet document.
            name (str): Name of sheet to move.
            idx (int): The zero based index to move sheet into.

        Returns:
            bool: ``True`` on success; Otherwise, ``False``.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_MOVING` :eventref:`src-docs-sheet-event-moving`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_MOVED` :eventref:`src-docs-sheet-event-moved`
        """
        cargs = SheetCancelArgs(Calc.move_sheet.__qualname__)
        cargs.doc = doc
        cargs.name = name
        cargs.index = idx
        _Events().trigger(CalcNamedEvent.SHEET_MOVING, cargs)
        if cargs.cancel:
            return False
        name = cargs.name
        idx = cargs.index
        sheets = cargs.doc.getSheets()
        num_sheets = len(sheets.getElementNames())
        result = False
        if idx < 0 or idx >= num_sheets:
            mLo.Lo.print(f"Index {idx} is out of range.")
        else:
            sheets.moveByName(name, idx)
            result = True
        if result:
            _Events().trigger(CalcNamedEvent.SHEET_MOVED, SheetArgs.from_args(cargs))
        return result

    # region get_sheet_names()
    @overload
    @classmethod
    def get_sheet_names(cls) -> Tuple[str, ...]:
        """
        Gets names of all existing spreadsheets in the spreadsheet document.

        |lo_unsafe|

        Args:
            doc (XSpreadsheetDocument, optional): Document to get sheets names of.

        Returns:
            Tuple[str, ...]: Tuple of sheet names.
        """
        ...

    @overload
    @classmethod
    def get_sheet_names(cls, doc: XSpreadsheetDocument) -> Tuple[str, ...]:
        """
        Gets names of all existing spreadsheets in the spreadsheet document.

        |lo_safe|

        Args:
            doc (XSpreadsheetDocument, optional): Document to get sheets names of.

        Returns:
            Tuple[str, ...]: Tuple of sheet names.
        """
        ...

    @classmethod
    def get_sheet_names(cls, doc: XSpreadsheetDocument | None = None) -> Tuple[str, ...]:
        """
        Gets names of all existing spreadsheets in the spreadsheet document.

        Args:
            doc (XSpreadsheetDocument, optional): Document to get sheets names of.

        Returns:
            Tuple[str, ...]: Tuple of sheet names.

        .. versionchanged:: 0.8.6
            Added overload ``get_sheet_names() -> Tuple[str, ...]``
        """
        if doc is None:
            doc = cls.get_current_doc()
        sheets = doc.getSheets()
        return sheets.getElementNames()

    # endregion get_sheet_names()

    # region get_sheets()
    @overload
    @classmethod
    def get_sheets(cls) -> XSpreadsheets:
        """
        Gets all existing spreadsheets in the spreadsheet document.

        |lo_unsafe|

        Returns:
            XSpreadsheets: document sheets
        """
        ...

    @overload
    @classmethod
    def get_sheets(cls, doc: XSpreadsheetDocument) -> XSpreadsheets:
        """
        Gets all existing spreadsheets in the spreadsheet document.

        |lo_safe|

        Args:
            doc (XSpreadsheetDocument, optional): Document to get sheets of.

        Returns:
            XSpreadsheets: document sheets
        """
        ...

    @classmethod
    def get_sheets(cls, doc: XSpreadsheetDocument | None = None) -> XSpreadsheets:
        """
        Gets all existing spreadsheets in the spreadsheet document.

        Args:
            doc (XSpreadsheetDocument, optional): Document to get sheets of.

        Returns:
            XSpreadsheets: document sheets
        """
        if doc is None:
            doc = cls.get_current_doc()
        return doc.getSheets()

    # endregion get_sheets()

    # region get_sheet_index()
    @overload
    @classmethod
    def get_sheet_index(cls) -> int:
        """
        Gets index if sheet.

        |lo_unsafe|

        Returns:
            int: Sheet Index.
        """
        ...

    @overload
    @classmethod
    def get_sheet_index(cls, sheet: XSpreadsheet) -> int:
        """
        Gets index if sheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spread sheet. Defaults to active sheet.

        Returns:
            int: Sheet Index.
        """
        ...

    @classmethod
    def get_sheet_index(cls, sheet: XSpreadsheet | None = None) -> int:
        """
        Gets index if sheet.

        Args:
            sheet (XSpreadsheet, optional): Spread sheet. Defaults to active sheet.

        Returns:
            int: Sheet Index.
        """
        if sheet is None:
            sheet = cls.get_active_sheet()
        ra = mLo.Lo.qi(XCellRangeAddressable, sheet, True)
        ca = ra.getRangeAddress()
        return ca.Sheet

    # endregion get_sheet_index()

    # region get_sheet_name()
    @overload
    @classmethod
    def get_sheet_name(cls) -> str:
        """
        Gets the name of a sheet.

        |lo_unsafe|

        Returns:
            str: Name of sheet
        """
        ...

    @overload
    @classmethod
    def get_sheet_name(cls, safe_quote: bool) -> str:
        """
        Gets the name of a sheet.

        |lo_unsafe|

        Args:
            safe_quote (bool, optional): If True, returns quoted (in single quotes) sheet name if the sheet name is not alphanumeric.
                Defaults to True.

        Returns:
            str: Name of sheet
        """
        ...

    @overload
    @classmethod
    def get_sheet_name(cls, idx: int) -> str:
        """
        Gets the name of a sheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet, optional): Spreadsheet
            idx (int, optional): Index of Spreadsheet
            safe_quote (bool, optional): If True, returns quoted (in single quotes) sheet name if the sheet name is not alphanumeric.
                Defaults to True.

        Returns:
            str: Name of sheet
        """
        ...

    @overload
    @classmethod
    def get_sheet_name(cls, idx: int, safe_quote: bool) -> str:
        """
        Gets the name of a sheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet, optional): Spreadsheet
            idx (int, optional): Index of Spreadsheet
            safe_quote (bool, optional): If True, returns quoted (in single quotes) sheet name if the sheet name is not alphanumeric.
                Defaults to True.

        Returns:
            str: Name of sheet
        """
        ...

    @overload
    @classmethod
    def get_sheet_name(cls, sheet: XSpreadsheet) -> str:
        """
        Gets the name of a sheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet, optional): Spreadsheet
            idx (int, optional): Index of Spreadsheet
            safe_quote (bool, optional): If True, returns quoted (in single quotes) sheet name if the sheet name is not alphanumeric.
                Defaults to True.

        Returns:
            str: Name of sheet
        """
        ...

    @overload
    @classmethod
    def get_sheet_name(cls, sheet: XSpreadsheet, safe_quote: bool) -> str:
        """
        Gets the name of a sheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet, optional): Spreadsheet
            idx (int, optional): Index of Spreadsheet
            safe_quote (bool, optional): If True, returns quoted (in single quotes) sheet name if the sheet name is not alphanumeric.
                Defaults to True.

        Returns:
            str: Name of sheet
        """
        ...

    @classmethod
    def get_sheet_name(cls, *args, **kwargs) -> str:
        """
        Gets the name of a sheet.

        |lo_unsafe|

        Args:
            sheet (XSpreadsheet, optional): Spreadsheet
            idx (int, optional): Index of Spreadsheet
            safe_quote (bool, optional): If True, returns quoted (in single quotes) sheet name if the sheet name is not alphanumeric.
                Defaults to True.

        Raises:
            MissingInterfaceError: If unable to access spreadsheet named interface

        Returns:
            str: Name of sheet

        .. versionchanged:: 0.8.6
            Added overload ``get_sheet_name(idx: int) -> str``

        .. versionchanged:: 0.11.12
            Added safe_quote parameter.
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = {"sheet", "idx", "safe_quote"}
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_sheet_name() got an unexpected keyword argument")
            for key in valid_keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            if "safe_quote" in kwargs:
                ka[2] = kwargs["safe_quote"]
            return ka

        if count not in (0, 1, 2):
            raise TypeError("get_sheet_name() got an invalid number of arguments")

        if count == 0:
            xnamed = mLo.Lo.qi(XNamed, cls.get_active_sheet(), True)
            return xnamed.getName()

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        safe_quote = True
        arg1 = kargs[1]  # "sheet", "idx", "safe_quote"
        if count == 1:
            if mInfo.Info.is_instance(arg1, bool):
                # int is not bool but bool is int
                safe_quote = arg1
                sheet = cls.get_active_sheet()
            elif mInfo.Info.is_instance(arg1, int):
                sheet = cls.get_sheet(arg1)
            else:
                sheet = cast(XSpreadsheet, arg1)
        else:
            # count == 2
            sheet = cast(XSpreadsheet, arg1)
            safe_quote = cast(bool, kargs[2])

        xnamed = mLo.Lo.qi(XNamed, sheet, True)
        sheet_name = xnamed.getName()
        # if sheet_name is not alphanumeric, then quote it
        if safe_quote and not sheet_name.isalnum():
            sheet_name = f"'{sheet_name}'"
        return sheet_name

    # endregion get_sheet_name()
    @overload
    @classmethod
    def set_sheet_name(cls, name: str) -> bool:
        """
        Sets the name of a spreadsheet.

        |lo_unsafe|

        Args:
            sheet (XSpreadsheet): Spreadsheet to set name of.
            name (str): New name for spreadsheet.

        Returns:
            bool: ``True`` on success; Otherwise, ``False``.
        """
        ...

    @overload
    @classmethod
    def set_sheet_name(cls, sheet: XSpreadsheet, name: str) -> bool:
        """
        Sets the name of a spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet to set name of.
            name (str): New name for spreadsheet.

        Returns:
            bool: ``True`` on success; Otherwise, ``False``.
        """
        ...

    @classmethod
    def set_sheet_name(cls, *args, **kwargs) -> bool:
        """
        Sets the name of a spreadsheet.

        Args:
            sheet (XSpreadsheet): Spreadsheet to set name of.
            name (str): New name for spreadsheet.

        Returns:
            bool: ``True`` on success; Otherwise, ``False``.
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("sheet", "name")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("set_sheet_name() got an unexpected keyword argument")
            keys = ("sheet", "name")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            ka[2] = kwargs.get("name")
            return ka

        if count not in (1, 2):
            raise TypeError("set_sheet_name() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            sheet = cls.get_active_sheet()
            name = cast(str, kargs[1])
        else:
            sheet = cast(XSpreadsheet, kargs[1])
            name = cast(str, kargs[2])

        xnamed = mLo.Lo.qi(XNamed, sheet)
        if xnamed is None:
            mLo.Lo.print("Could not access spreadsheet")
            return False
        xnamed.setName(name)
        return True

    @staticmethod
    def protect_sheet(sheet: XSpreadsheet, password: str) -> bool:
        """
        Protects a Spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Sheet to protect.
            password (str): Password to protect sheet with.

        Returns:
            bool: ``True`` on success; Otherwise, ``False``

        See Also:
            - :py:meth:`~.calc.Calc.unprotect_sheet`
            - :py:meth:`~.calc.Calc.is_sheet_protected`
            - :ref:`help_calc_format_direct_cell_cell_protection`

        .. versionadded:: 0.10.0
        """
        pro = mLo.Lo.qi(XProtectable, sheet, True)
        if not pro.isProtected():
            pro.protect(password)
            return True
        return False

    @staticmethod
    def unprotect_sheet(sheet: XSpreadsheet, password: str) -> bool:
        """
        Unprotect a Spreadsheet.

        If sheet is not protected, this method will still return ``True``.

        If incorrect password is provided, this method will return ``False``.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Sheet to unprotect.
            password (str): Password to unprotect sheet with.

        Returns:
            bool: ``True`` on success; Otherwise, ``False``

        See Also:
            - :py:meth:`~.calc.Calc.protect_sheet`
            - :py:meth:`~.calc.Calc.is_sheet_protected`
            - :ref:`help_calc_format_direct_cell_cell_protection`

        .. versionadded:: 0.10.0
        """
        pro = mLo.Lo.qi(XProtectable, sheet, True)
        if pro.isProtected():
            try:
                pro.unprotect(password)
            except IllegalArgumentException:
                return False
        return pro.isProtected()

    @staticmethod
    def is_sheet_protected(sheet: XSpreadsheet) -> bool:
        """
        Gets whether a sheet is protected.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Sheet to check

        Returns:
            bool: True if protected; Otherwise, False

        See Also:
            - :py:meth:`~.calc.Calc.protect_sheet`
            - :py:meth:`~.calc.Calc.unprotect_sheet`
            - :ref:`help_calc_format_direct_cell_cell_protection`

        .. versionadded:: 0.10.0
        """
        pro = mLo.Lo.qi(XProtectable, sheet, True)
        return pro.isProtected()

    # endregion --------------------- sheet methods -------------------------

    # region --------------- view methods ------------------------------

    # region get_controller()
    @overload
    @classmethod
    def get_controller(cls) -> XController: ...

    @overload
    @classmethod
    def get_controller(cls, doc: XSpreadsheetDocument) -> XController: ...

    @classmethod
    def get_controller(cls, doc: XSpreadsheetDocument | None = None) -> XController:
        """
        Provides access to the controller which currently controls this model.

        |lo_unsafe|

        Args:
            doc (XSpreadsheetDocument, optional): Spreadsheet Document

        Raises:
            MissingInterfaceError: If unable to access controller

        Returns:
            XController | None: Controller for Spreadsheet Document
        """
        if doc is None:
            doc = cls.get_current_doc()
        model = mLo.Lo.qi(XModel, doc, True)
        return model.getCurrentController()

    # endregion get_controller()

    @classmethod
    def zoom_value(cls, doc: XSpreadsheetDocument, value: int) -> None:
        """
        Sets the zoom level of the Spreadsheet Document.

        |lo_unsafe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document
            value (int): Value to set zoom. e.g. 160 set zoom to 160%
        """
        ctrl = cls.get_controller(doc)
        if ctrl is None:
            return
        mProps.Props.set(ctrl, ZoomType=int(ZoomKind.BY_VALUE), ZoomValue=value)

    @classmethod
    def zoom(cls, doc: XSpreadsheetDocument, type: ZoomKind) -> None:
        """
        Zooms spreadsheet document to a specific view.

        |lo_unsafe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document
            type (GUI.ZoomEnum): Type of Zoom to set.
        """

        ctrl = cls.get_controller(doc)
        if ctrl is None:
            return

        def zoom_val(value: int) -> None:
            mProps.Props.set(ctrl, ZoomType=int(ZoomKind.BY_VALUE), ZoomValue=value)

        if type in (
            ZoomKind.ENTIRE_PAGE,
            ZoomKind.OPTIMAL,
            ZoomKind.PAGE_WIDTH,
            ZoomKind.PAGE_WIDTH_EXACT,
        ):
            mProps.Props.set(ctrl, ZoomType=type.value)
        elif type == ZoomKind.ZOOM_200_PERCENT:
            zoom_val(200)
        elif type == ZoomKind.ZOOM_150_PERCENT:
            zoom_val(150)
        elif type == ZoomKind.ZOOM_100_PERCENT:
            zoom_val(100)
        elif type == ZoomKind.ZOOM_75_PERCENT:
            zoom_val(75)
        elif type == ZoomKind.ZOOM_50_PERCENT:
            zoom_val(50)

    @classmethod
    def get_view(cls, doc: XSpreadsheetDocument) -> XSpreadsheetView:
        """
        Is the main interface of a SpreadsheetView.

        It manages the active sheet within this view.

        The ``com.sun.star.sheet.SpreadsheetView`` service is the spreadsheet's extension
        of the ``com.sun.star.frame.Controller`` service and represents a table editing view
        for a spreadsheet document.

        |lo_unsafe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document.

        Returns:
            XSpreadsheetView: ``XSpreadsheetView``.
        """
        return mLo.Lo.qi(XSpreadsheetView, cls.get_controller(doc), True)

    @classmethod
    def set_active_sheet(cls, doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> None:
        """
        Sets the active sheet.

        |lo_unsafe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document
            sheet (XSpreadsheet): Sheet to set active

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_ACTIVATING` :eventref:`src-docs-sheet-event-activating`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_ACTIVATED` :eventref:`src-docs-sheet-event-activated`

        Note:
            Event arg properties modified on SHEET_ACTIVATING it is reflected in this method.
        """
        cargs = SheetCancelArgs(Calc.set_active_sheet.__qualname__)
        cargs.doc = doc
        cargs.sheet = sheet
        _Events().trigger(CalcNamedEvent.SHEET_ACTIVATING, cargs)
        if cargs.cancel:
            return
        ss_view = cls.get_view(cargs.doc)
        if ss_view is None:
            return
        ss_view.setActiveSheet(cargs.sheet)
        _Events().trigger(CalcNamedEvent.SHEET_ACTIVATED, SheetArgs.from_args(cargs))

    # region get_active_sheet()
    @overload
    @classmethod
    def get_active_sheet(cls) -> XSpreadsheet: ...

    @overload
    @classmethod
    def get_active_sheet(cls, doc: XSpreadsheetDocument) -> XSpreadsheet: ...

    @classmethod
    def get_active_sheet(cls, doc: XSpreadsheetDocument | None = None) -> XSpreadsheet:
        """
        Gets the active sheet.

        |lo_unsafe|

        Args:
            doc (XSpreadsheetDocument, optional): Spreadsheet Document.

        Returns:
            XSpreadsheet: Active Sheet.
        """
        if doc is None:
            doc = cls.get_current_doc()
        ss_view = cls.get_view(doc)
        return ss_view.getActiveSheet()

    # endregion get_active_sheet()

    @classmethod
    def freeze(cls, doc: XSpreadsheetDocument, num_cols: int, num_rows: int) -> None:
        """
        Freezes spreadsheet columns and rows.

        |lo_unsafe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document.
            num_cols (int): Number of columns to freeze.
            num_rows (int): Number of rows to freeze.

        Returns:
            None:

        See Also:

            - :ref:`ch23_freezing_rows`
            - :py:meth:`~.Calc.freeze_rows`
            - :py:meth:`~.Calc.freeze_cols`
            - :py:meth:`~.Calc.unfreeze`
        """
        ctrl = cls.get_controller(doc)
        if ctrl is None:
            return
        if num_cols < 0 or num_rows < 0:
            return
        freeze = mLo.Lo.qi(XViewFreezable, ctrl, True)
        freeze.freezeAtPosition(num_cols, num_rows)

    @classmethod
    def unfreeze(cls, doc: XSpreadsheetDocument) -> None:
        """
        UN-Freezes spreadsheet columns and/or rows.

        |lo_unsafe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document.

        Returns:
            None:

        See Also:

            - :ref:`ch23_freezing_rows`
            - :py:meth:`~.Calc.freeze`
            - :py:meth:`~.Calc.freeze_rows`
            - :py:meth:`~.Calc.freeze_cols`
        """
        ctrl = cls.get_controller(doc)
        if ctrl is None:
            return
        freeze = mLo.Lo.qi(XViewFreezable, ctrl, True)
        if freeze.hasFrozenPanes():
            cls.freeze(doc=doc, num_cols=0, num_rows=0)

    @classmethod
    def freeze_cols(cls, doc: XSpreadsheetDocument, num_cols: int) -> None:
        """
        Freezes spreadsheet columns.

        |lo_unsafe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document.
            num_cols (int): Number of columns to freeze.

        Returns:
            None:

        See Also:

            - :ref:`ch23_freezing_rows`
            - :py:meth:`~.Calc.freeze`
            - :py:meth:`~.Calc.freeze_rows`
            - :py:meth:`~.Calc.unfreeze`
        """
        cls.freeze(doc=doc, num_cols=num_cols, num_rows=0)

    @classmethod
    def freeze_rows(cls, doc: XSpreadsheetDocument, num_rows: int) -> None:
        """
        Freezes spreadsheet rows.

        |lo_unsafe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document.
            num_rows (int): Number of rows to freeze.

        Returns:
            None:

        See Also:

            - :ref:`ch23_freezing_rows`
            - :py:meth:`~.Calc.freeze`
            - :py:meth:`~.Calc.freeze_cols`
            - :py:meth:`~.Calc.unfreeze`
        """
        cls.freeze(doc=doc, num_cols=0, num_rows=num_rows)

    # region    goto_cell()
    @overload
    @classmethod
    def goto_cell(cls, cell_name: str, doc: XSpreadsheetDocument) -> None: ...

    @overload
    @classmethod
    def goto_cell(cls, cell_obj: mCellObj.CellObj, doc: XSpreadsheetDocument) -> None: ...

    @overload
    @classmethod
    def goto_cell(cls, cell_name: str, frame: XFrame) -> None: ...

    @overload
    @classmethod
    def goto_cell(cls, cell_obj: mCellObj.CellObj, frame: XFrame) -> None: ...

    @classmethod
    def goto_cell(cls, *args, **kwargs) -> None:
        """
        Go to a cell.

        |lo_unsafe|

        Args:
            cell_name (str): Cell Name such as 'B4'
            doc (XSpreadsheetDocument): Spreadsheet Document
            frame (XFrame): Spreadsheet frame.

        Attention:
            :py:meth:`~.utils.lo.Lo.dispatch_cmd` method is called along with any of its events.

            Dispatch command is ``GoToCell``.
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs():
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("cell_name", "cell_obj", "doc", "frame")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("goto_cell() got an unexpected keyword argument")
            keys = ("cell_name", "cell_obj")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            keys = ("doc", "frame")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            return ka

        if count != 2:
            raise TypeError("set_val() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        doc = mLo.Lo.qi(XSpreadsheetDocument, kargs[2])
        frame = cls.get_controller(doc).getFrame() if doc is not None else kargs[2]
        props = mProps.Props.make_props(ToPoint=str(kargs[1]))
        mLo.Lo.dispatch_cmd(cmd="GoToCell", props=props, frame=frame)

    # endregion    goto_cell()

    @classmethod
    def split_window(cls, doc: XSpreadsheetDocument, cell_name: str) -> None:
        """
        Splits window.

        |lo_unsafe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document.
            cell_name (str): Cell to preform split on. e.g. ``C4``.

        Returns:
            None:

        See Also:
            :ref:`ch23_splitting_panes`
        """
        frame = cls.get_controller(doc).getFrame()
        cls.goto_cell(cell_name=cell_name, frame=frame)
        props = mProps.Props.make_props(ToPoint=cell_name)
        mLo.Lo.dispatch_cmd(cmd="SplitWindow", props=props, frame=frame)

    # region    get_selected_addr()

    @overload
    @classmethod
    def get_selected_addr(cls, doc: XSpreadsheetDocument) -> CellRangeAddress:
        """
        Gets select cell range addresses.

        |lo_safe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document.

        Returns:
            CellRangeAddress: Cell range addresses.
        """
        ...

    @overload
    @classmethod
    def get_selected_addr(cls, model: XModel) -> CellRangeAddress:
        """
        Gets select cell range addresses.

        |lo_safe|

        Args:
            model (XModel): model used to access sheet.

        Returns:
            CellRangeAddress: Cell range addresses.
        """
        ...

    @classmethod
    def get_selected_addr(cls, *args, **kwargs) -> CellRangeAddress:
        """
        Gets select cell range addresses.

        |lo_safe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document.
            model (XModel): model used to access sheet.

        Raises:
            Exception: if unable to get document model.
            MissingInterfaceError: if unable to get interface ``XCellRangeAddressable``.

        Returns:
            CellRangeAddress: Cell range addresses.

        See Also:
            - :py:meth:`~.Calc.get_selected_range`
            - :py:meth:`~.Calc.set_selected_addr`
            - :py:meth:`~.Calc.set_selected_range`
            - :py:meth:`~.Calc.get_selected_cell_addr`
        """
        # sourcery skip: raise-specific-error
        ordered_keys = (1,)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("doc", "model")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_selected_addr() got an unexpected keyword argument")
            keys = ("doc", "model")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            return ka

        if count != 1:
            raise TypeError("get_selected_addr() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        doc = mLo.Lo.qi(XSpreadsheetDocument, kargs[1])
        model = mLo.Lo.qi(XModel, doc) if doc is not None else cast(XModel, kargs[1])
        if model is None:
            raise Exception("No document model found")
        ra = mLo.Lo.qi(XCellRangeAddressable, model.getCurrentSelection(), raise_err=True)
        return ra.getRangeAddress()  # type: ignore

    # endregion  get_selected_addr()

    # region get_selected_range()
    @overload
    @classmethod
    def get_selected_range(cls, doc: XSpreadsheetDocument) -> mRngObj.RangeObj: ...

    @overload
    @classmethod
    def get_selected_range(cls, model: XModel) -> mRngObj.RangeObj: ...

    @classmethod
    def get_selected_range(cls, *args, **kwargs) -> mRngObj.RangeObj:
        """
        Gets select cell range.

        |lo_safe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document
            model (XModel): model used to access sheet

        Raises:
            Exception: if unable to get document model
            MissingInterfaceError: if unable to get interface XCellRangeAddressable

        Returns:
            RangeObj: Cell range addresses

        See Also:
            - :py:meth:`~.Calc.get_selected_addr`
            - :py:meth:`~.Calc.set_selected_addr`
            - :py:meth:`~.Calc.get_selected_cell_addr`
            - :py:meth:`~.Calc.set_selected_range`

        .. versionadded:: 0.8.2
        """
        ca = cls.get_selected_addr(*args, **kwargs)
        return cls.get_range_obj(ca)

    # endregion get_selected_range()

    @classmethod
    def get_selected_cell_addr(cls, doc: XSpreadsheetDocument) -> CellAddress:
        """
        Gets the cell address of current selected cell of the active sheet.

        |lo_unsafe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet document

        Raises:
            CellError: if active selection is not a single cell

        Returns:
            CellAddress: Cell Address

        Note:
            CellAddress returns Zero-base values.
            For instance: Cell ``B4`` has Column value of ``1`` and Row value of ``3``

        See Also:
            - :py:meth:`~.Calc.get_selected_cell`
            - :py:meth:`~.Calc.get_selected_addr`
            - :py:meth:`~.Calc.set_selected_addr`
        """
        cr_addr = cls.get_selected_addr(doc=doc)
        if cls.is_single_cell_range(cr_addr):
            sheet = cls.get_active_sheet(doc)
            cell = cls.get_cell(sheet=sheet, col=cr_addr.StartColumn, row=cr_addr.StartRow)
            return cls.get_cell_address(cell)
        else:
            raise mEx.CellError("Selected address is not a single cell")

    # region get_selected_cell()

    @overload
    @classmethod
    def get_selected_cell(cls) -> mCellObj.CellObj: ...

    @overload
    @classmethod
    def get_selected_cell(cls, doc: XSpreadsheetDocument) -> mCellObj.CellObj: ...

    @classmethod
    def get_selected_cell(cls, doc: XSpreadsheetDocument | None = None) -> mCellObj.CellObj:
        """
        Gets the cell address of current selected cell of the active sheet.

        |lo_unsafe|

        Args:
            doc (XSpreadsheetDocument, Optional): Spreadsheet document

        Raises:
            CellError: if active selection is not a single cell

        Returns:
            CellAddress: Cell Address

        Note:
            CellAddress returns Zero-base values.
            For instance: Cell ``B4`` has Column value of ``1`` and Row value of ``3``

        See Also:
            - :py:meth:`~.Calc.get_selected_cell_addr`
            - :py:meth:`~.Calc.get_selected_addr`
            - :py:meth:`~.Calc.set_selected_addr`
        """
        if doc is None:
            doc = cls.get_current_doc()
        ca = cls.get_selected_cell_addr(doc)
        return mCellObj.CellObj.from_idx(col_idx=ca.Column, row_idx=ca.Row)

    # endregion get_selected_cell()

    # region set_selected_addr()

    @overload
    @classmethod
    def set_selected_addr(cls, doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> CellRangeAddress | None:
        """
        Selects cells in a Spreadsheet.

        If ``range_name`` is omitted then deselection is preformed.

        |lo_unsafe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet document.
            sheet (XSpreadsheet): Spreadsheet.
            range_val (str | RangeObj): Range name.

        Returns:
            CellRangeAddress | None: Cell range address of the current selection if successful, otherwise ``None``
        """
        ...

    @overload
    @classmethod
    def set_selected_addr(
        cls, doc: XSpreadsheetDocument, sheet: XSpreadsheet, range_val: str | mRngObj.RangeObj
    ) -> CellRangeAddress | None:
        """
        Selects cells in a Spreadsheet.

        If ``range_name`` is omitted then deselection is preformed.

        |lo_safe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet document.
            sheet (XSpreadsheet): Spreadsheet.
            range_val (str | RangeObj): Range name.

        Returns:
            CellRangeAddress | None: Cell range address of the current selection if successful, otherwise ``None``
        """
        ...

    @classmethod
    def set_selected_addr(cls, *args, **kwargs) -> CellRangeAddress | None:
        """
        Selects cells in a Spreadsheet.

        If ``range_name`` is omitted then deselection is preformed.

        Args:
            doc (XSpreadsheetDocument): Spreadsheet document
            sheet (XSpreadsheet): Spreadsheet
            range_val (str | RangeObj): Range name

        Returns:
            CellRangeAddress | None: Cell range address of the current selection if successful, otherwise ``None``

        See Also:
            - :py:meth:`~.Calc.get_selected_addr`
            - :py:meth:`~.Calc.get_selected_cell_addr`

        .. versionadded:: 0.8.1
        """
        # range_name is backwards compatibility, Changed in ver 0.8.3
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("doc", "sheet", "range_name", "range_val")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("set_selected_addr() got an unexpected keyword argument")
            ka[1] = kwargs.get("doc", None)
            ka[2] = kwargs.get("sheet", None)
            keys = ("range_name", "range_val")
            for key in keys:
                if key in kwargs:
                    ka[3] = kwargs[key]
                    break
            return ka

        if count not in (2, 3):
            raise TypeError("set_selected_addr() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        # this method works fine in headless mode.
        doc = kargs[1]
        sheet = cast(XSpreadsheet, kargs[2])
        if count == 2:
            sel_obj = mLo.Lo.create_instance_msf(XCellRangesQuery, "com.sun.star.sheet.SheetCellRanges")
            if sel_obj is None:
                return None
        else:
            sel_obj = sheet.getCellRangeByName(str(kargs[3]))
        supp = mGui.GUI.get_selection_supplier(doc)
        supp.select(sel_obj)
        return cls.get_selected_addr(doc)

    # endregion set_selected_addr()

    # region set_selected()
    @overload
    @classmethod
    def set_selected_range(
        cls, doc: XSpreadsheetDocument, sheet: XSpreadsheet, range_val: mRngObj.RangeObj
    ) -> mRngObj.RangeObj:
        """
        Selects cells in a Spreadsheet.

        If ``range_name`` is omitted then deselection is preformed.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_val (str): Range name such as ``A1:G3`` or ``RangeObj``.

        Returns:
            CellRangeAddress | None: Cell range address of the current selection if successful, otherwise ``None``.
        """
        ...

    @overload
    @classmethod
    def set_selected_range(cls, doc: XSpreadsheetDocument, sheet: XSpreadsheet, range_val: str) -> mRngObj.RangeObj:
        """
        Selects cells in a Spreadsheet.

        If ``range_name`` is omitted then deselection is preformed.

        |lo_safe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet document.
            sheet (XSpreadsheet): Spreadsheet.
            range_val (str): Range name such as ``A1:G3`` or ``RangeObj``.

        Returns:
            CellRangeAddress | None: Cell range address of the current selection if successful, otherwise ``None``.
        """
        ...

    @overload
    @classmethod
    def set_selected_range(cls, doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> mRngObj.RangeObj:
        """
        Selects cells in a Spreadsheet.

        If ``range_name`` is omitted then deselection is preformed.

        |lo_safe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet document.
            sheet (XSpreadsheet): Spreadsheet.

        Returns:
            CellRangeAddress | None: Cell range address of the current selection if successful, otherwise ``None``.
        """
        ...

    @classmethod
    def set_selected_range(
        cls, doc: XSpreadsheetDocument, sheet: XSpreadsheet, range_val: str | mRngObj.RangeObj = ""
    ) -> mRngObj.RangeObj | None:
        """
        Selects cells in a Spreadsheet.

        If ``range_name`` is omitted then deselection is preformed.

        |lo_safe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet document.
            sheet (XSpreadsheet): Spreadsheet.
            range_val (str): Range name such as ``A1:G3`` or ``RangeObj``.

        Returns:
            CellRangeAddress | None: Cell range address of the current selection if successful, otherwise ``None``.

        See Also:
            - :py:meth:`~.Calc.get_selected_range`
            - :py:meth:`~.Calc.get_selected_addr`
            - :py:meth:`~.Calc.get_selected_cell_addr`
            - :py:meth:`~.Calc.set_selected_addr`

        .. versionadded:: 0.8.2
        """
        rng_name = str(range_val) if range_val else ""
        ca = cls.set_selected_addr(doc=doc, sheet=sheet, range_val=rng_name)
        return None if ca is None else cls.get_range_obj(ca)

    # endregion set_selected()

    # endregion -------------- view methods ----------------------------

    # region --------------- view data methods -------------------------

    @classmethod
    def get_view_panes(cls, doc: XSpreadsheetDocument) -> List[XViewPane] | None:
        """
        Represents a pane in a view of a spreadsheet document.

        |lo_unsafe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document.

        Raises:
            MissingInterfaceError: if unable access the view pane container.

        Returns:
            List[XViewPane] | None: List of XViewPane on success; Otherwise, None.

        Note:
            The com.sun.star.sheet.XViewPane interface's getFirstVisibleColumn(), getFirstVisibleRow(),
            setFirstVisibleColumn() and setFirstVisibleRow() methods query and set the start of
            the exposed area. The getVisibleRange() method returns a com.sun.star.table.
            CellRangeAddress struct describing which cells are shown in the pane.
            Columns or rows that are only partly visible at the right or lower edge of the view
            are not included.

        See Also:
            :ref:`ch23_view_states_top_pane`
        """
        con = mLo.Lo.qi(XIndexAccess, cls.get_controller(doc), True)

        panes = []
        for i in range(con.getCount()):
            try:
                panes.append(mLo.Lo.qi(XViewPane, con.getByIndex(i)))
            except UnoException:
                mLo.Lo.print(f"Could not get view pane {i}")
        if not panes:
            mLo.Lo.print("No view panes found")
            return None
        return panes

    @classmethod
    def get_view_data(cls, doc: XSpreadsheetDocument) -> str:
        """
        Gets a set of data that can be used to restore the current view status at
        later time by using :py:meth:`~Calc.set_view_data`.

        |lo_unsafe|

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

        |lo_unsafe|

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

        The view data string has the format
        ``100/60/0;0;tw:879;0/4998/0/1/0/218/2/0/0/4988/4998``.

        The view state info starts after the third ``;``, the fourth entry.
        The view state for each sheet is separated by ``;``

        Based on a post by user Hanya to:
        `openoffice forum <https://forum.openoffice.org/en/forum/viewtopic.php?f=45&t=29195&p=133202&hilit=getViewData#p133202>`_

        |lo_unsafe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document

        Returns:
            None:

        See Also:
            :ref:`ch23_view_states_top_pane`
        """
        ctrl = cls.get_controller(doc)

        view_data = str(ctrl.getViewData())
        view_parts = view_data.split(";")
        p_len = len(view_parts)
        if p_len < 4:
            mLo.Lo.print("No sheet view states found in view data")
            return None
        return [mViewState.ViewState(view_parts[i]) for i in range(3, p_len)]

    @classmethod
    def set_view_states(cls, doc: XSpreadsheetDocument, states: Sequence[mViewState.ViewState]) -> None:
        """
        Updates the sheet state part of the view data, which starts as the fourth entry in the view data string.

        |lo_unsafe|

        Args:
            doc (XSpreadsheetDocument): spreadsheet document
            states (Sequence[ViewState]): Sequence of ViewState objects.

        Returns:
            None:

        See Also:
            :ref:`ch23_view_states_top_pane`
        """
        ctrl = cls.get_controller(doc)
        if ctrl is None:
            return
        view_data = str(ctrl.getViewData())
        view_parts = view_data.split(";")
        p_len = len(view_parts)
        if p_len < 4:
            mLo.Lo.print("No sheet view states found in view data")
            return None

        vd_new = [view_parts[i] for i in range(3)]
        vd_new.extend(str(state) for state in states)
        s_data = ";".join(vd_new)
        mLo.Lo.print(s_data)
        ctrl.restoreViewData(s_data)

    # endregion ----------------- view data methods ---------------------------------

    # region --------------- insert/remove rows, columns, cells --------

    @staticmethod
    def insert_row(sheet: XSpreadsheet, idx: int, count: int = 1) -> bool:
        """
        Inserts a row in spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet
            idx (int): Zero base index of row to insert.
            count (int, optional): Number of rows to insert. Defaults to ``1``.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_ROW_INSERTING` :eventref:`src-docs-sheet-event-row-inserting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_INSERTED` :eventref:`src-docs-sheet-event-row-inserted`

        Returns:
            bool: ``True`` if row has been inserted; Otherwise, ``False``.

        .. versionchanged:: 0.15.0
            Added ``count`` parameter
        """
        cargs = SheetCancelArgs(Calc.insert_row.__qualname__)
        cargs.sheet = sheet
        cargs.index = idx
        _Events().trigger(CalcNamedEvent.SHEET_ROW_INSERTING, cargs)
        if cargs.cancel:
            return False
        idx = cargs.index
        cr_range = mLo.Lo.qi(XColumnRowRange, cargs.sheet, True)
        rows = cr_range.getRows()
        rows.insertByIndex(idx, count)  # add 1 row at idx position
        _Events().trigger(CalcNamedEvent.SHEET_ROW_INSERTED, SheetArgs.from_args(cargs))
        return True

    @staticmethod
    def delete_row(sheet: XSpreadsheet, idx: int, count: int = 1) -> bool:
        """
        Deletes a row from spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            idx (int): Zero based index of row to delete.
            count (int, optional): Number of rows to delete. Defaults to ``1``.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_ROW_DELETING` :eventref:`src-docs-sheet-event-row-deleting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_ROW_DELETED` :eventref:`src-docs-sheet-event-row-deleted`

        Returns:
            bool: ``True`` if row is deleted; Otherwise, ``False``.

        .. versionchanged:: 0.15.0
            Added ``count`` parameter
        """
        cargs = SheetCancelArgs(Calc.delete_row.__qualname__)
        cargs.sheet = sheet
        cargs.index = idx
        cargs.name = None
        _Events().trigger(CalcNamedEvent.SHEET_ROW_DELETING, cargs)
        if cargs.cancel:
            return False
        idx = cargs.index
        cr_range = mLo.Lo.qi(XColumnRowRange, cargs.sheet, True)
        rows = cr_range.getRows()
        rows.removeByIndex(idx, count)  # remove row(s) at idx position
        _Events().trigger(CalcNamedEvent.SHEET_ROW_DELETED, SheetArgs.from_args(cargs))
        return True

    @staticmethod
    def insert_column(sheet: XSpreadsheet, idx: int, count: int = 1) -> bool:
        """
        Inserts a column in a spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            idx (int): Zero base index of column to insert.
            count (int, optional): Number of columns to insert. Defaults to ``1``.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_COL_INSERTING` :eventref:`src-docs-sheet-event-col-inserting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_COL_INSERTED` :eventref:`src-docs-sheet-event-col-inserted`

        Returns:
            bool: ``True`` if column has been inserted; Otherwise, ``False``.

        .. versionchanged:: 0.15.0
            Added ``count`` parameter
        """
        cargs = SheetCancelArgs(Calc.insert_column.__qualname__)
        cargs.sheet = sheet
        cargs.index = idx
        _Events().trigger(CalcNamedEvent.SHEET_COL_INSERTING, cargs)
        if cargs.cancel:
            return False
        idx = cargs.index
        cr_range = mLo.Lo.qi(XColumnRowRange, cargs.sheet, True)
        cols = cr_range.getColumns()
        cols.insertByIndex(idx, count)  # add column(s) at idx position
        _Events().trigger(CalcNamedEvent.SHEET_COL_INSERTED, SheetArgs.from_args(cargs))
        return True

    @staticmethod
    def delete_column(sheet: XSpreadsheet, idx: int, count: int = 1) -> bool:
        """
        Delete a column from a spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            idx (int): Zero base of index of column to delete.
            count (int, optional): Number of columns to delete. Defaults to ``1``.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_COL_DELETING` :eventref:`src-docs-sheet-event-col-deleting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_COL_DELETED` :eventref:`src-docs-sheet-event-col-deleted`

        Returns:
            bool: ``True`` if column is deleted; Otherwise, ``False``.

        .. versionchanged:: 0.15.0
            Added ``count`` parameter
        """
        cargs = SheetCancelArgs(Calc.delete_column.__qualname__)
        cargs.sheet = sheet
        cargs.index = idx
        _Events().trigger(CalcNamedEvent.SHEET_COL_DELETING, cargs)
        if cargs.cancel:
            return False
        idx = cargs.index
        cr_range = mLo.Lo.qi(XColumnRowRange, cargs.sheet, True)
        cols = cr_range.getColumns()
        cols.removeByIndex(idx, count)  # remove column(s) at idx position
        _Events().trigger(CalcNamedEvent.SHEET_COL_DELETED, SheetArgs.from_args(cargs))
        return True

    # region insert_cells()
    @classmethod
    def _insert_cells(cls, sheet: XSpreadsheet, cell_range: XCellRange, is_shift_right: bool) -> bool:
        """LO Safe Method."""
        cargs = CellCancelArgs(Calc.insert_cells.__qualname__)
        cargs.sheet = sheet
        cargs.cells = cell_range
        cargs.event_data = {"is_shift_right": is_shift_right}
        _Events().trigger(CalcNamedEvent.CELLS_INSERTING, cargs)
        if cargs.cancel:
            return False
        mover = mLo.Lo.qi(XCellRangeMovement, cargs.sheet, True)
        addr = cls.get_address(cargs.cells)
        if cargs.event_data["is_shift_right"]:
            mover.insertCells(addr, CellInsertMode.RIGHT)  # type: ignore
        else:
            mover.insertCells(addr, CellInsertMode.DOWN)  # type: ignore
        _Events().trigger(CalcNamedEvent.CELLS_INSERTED, CellArgs.from_args(cargs))
        return True

    @overload
    @classmethod
    def insert_cells(cls, sheet: XSpreadsheet, cell_range: XCellRange, is_shift_right: bool) -> bool:
        """
        Inserts Cells into a spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_range (XCellRange): Cell range to insert.
            is_shift_right (bool): If True then cell are inserted to the right; Otherwise, inserted down.

        Returns:
            bool: ``True`` if cells are inserted; Otherwise, ``False``.
        """
        ...

    @overload
    @classmethod
    def insert_cells(cls, sheet: XSpreadsheet, cr_addr: CellRangeAddress, is_shift_right: bool) -> bool:
        """
        Inserts Cells into a spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cr_addr (CellRangeAddress): Cell range Address.
            is_shift_right (bool): If True then cell are inserted to the right; Otherwise, inserted down.

        Returns:
            bool: ``True`` if cells are inserted; Otherwise, ``False``.
        """
        ...

    @overload
    @classmethod
    def insert_cells(cls, sheet: XSpreadsheet, range_name: str, is_shift_right: bool) -> bool:
        """
        Inserts Cells into a spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range Name such as ``A1:D5``.
            is_shift_right (bool): If True then cell are inserted to the right; Otherwise, inserted down.

        Returns:
            bool: ``True`` if cells are inserted; Otherwise, ``False``.
        """
        ...

    @overload
    @classmethod
    def insert_cells(cls, sheet: XSpreadsheet, range_obj: mRngObj.RangeObj, is_shift_right: bool) -> bool:
        """
        Inserts Cells into a spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_obj (RangeObj): Range Object.
            is_shift_right (bool): If True then cell are inserted to the right; Otherwise, inserted down.

        Returns:
            bool: ``True`` if cells are inserted; Otherwise, ``False``.
        """
        ...

    @overload
    @classmethod
    def insert_cells(
        cls, sheet: XSpreadsheet, col_start: int, row_start: int, col_end: int, row_end: int, is_shift_right: bool
    ) -> bool:
        """
        Inserts Cells into a spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.
            is_shift_right (bool): If True then cell are inserted to the right; Otherwise, inserted down.

        Returns:
            bool: ``True`` if cells are inserted; Otherwise, ``False``.
        """
        ...

    @classmethod
    def insert_cells(cls, *args, **kwargs) -> bool:
        """
        Inserts Cells into a spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_range (XCellRange): Cell range to insert.
            cr_addr (CellRangeAddress): Cell range Address.
            range_name (str): Range Name such as ``A1:D5``.
            range_obj (RangeObj): Range Object.
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.
            is_shift_right (bool): If True then cell are inserted to the right; Otherwise, inserted down.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_INSERTING` :eventref:`src-docs-cell-event-inserting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_INSERTED` :eventref:`src-docs-cell-event-inserted`

        Returns:
            bool: ``True`` if cells are inserted; Otherwise, ``False``.

        Note:
            Events args for this method have a ``cell`` type of ``XCellRange``

            Event args ``event_data`` is a dictionary containing ``is_shift_right``.
        """
        kw = kwargs.copy()
        sheet = kw.get("sheet", None)
        lst_args = list(args)
        if sheet is None:
            sheet = lst_args[0]

        is_shift_right = kw.get("is_shift_right", None)
        # is_shift_left needs to be removed to pass args to get_cell_range()
        if is_shift_right is None:
            is_shift_right = bool(lst_args.pop())
        else:
            del kw["is_shift_right"]

        cell_range = None
        arg2 = kwargs.get("cell_range", None)
        if arg2 is None and len(lst_args) > 1:
            arg2 = lst_args[1]

        if arg2:
            cell_range = mLo.Lo.qi(XCellRange, arg2)

        if cell_range:
            return cls._insert_cells(sheet=sheet, cell_range=cell_range, is_shift_right=is_shift_right)

        cell_range = cls.get_cell_range(*lst_args, **kw)
        return cls._insert_cells(sheet=sheet, cell_range=cell_range, is_shift_right=is_shift_right)

    # endregion insert_cells()

    # region delete_cells()

    @classmethod
    def _delete_cells(cls, sheet: XSpreadsheet, cell_range: XCellRange, is_shift_left: bool) -> bool:
        """LO Safe Method."""
        cargs = CellCancelArgs(Calc.delete_cells.__qualname__)
        cargs.sheet = sheet
        cargs.cells = cell_range
        cargs.event_data = {"is_shift_left": is_shift_left}
        _Events().trigger(CalcNamedEvent.CELLS_DELETING, cargs)
        if cargs.cancel:
            return False
        mover = mLo.Lo.qi(XCellRangeMovement, cargs.sheet)
        addr = cls.get_address(cargs.cells)
        if cargs.event_data["is_shift_left"]:
            mover.removeRange(addr, CellDeleteMode.LEFT)  # type: ignore
        else:
            mover.removeRange(addr, CellDeleteMode.UP)  # type: ignore
        _Events().trigger(CalcNamedEvent.CELLS_DELETED, CellArgs.from_args(cargs))
        return True

    @overload
    @classmethod
    def delete_cells(cls, sheet: XSpreadsheet, cell_range: XCellRange, is_shift_left: bool) -> bool:
        """
        Deletes cells in a spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_range (XCellRange): Cell range to delete.
            is_shift_left (bool): If ``True`` then cell are shifted left; Otherwise, cells are shifted up.

        Returns:
            bool: ``True`` if cells are deleted; Otherwise, ``False``.
        """
        ...

    @overload
    @classmethod
    def delete_cells(cls, sheet: XSpreadsheet, cr_addr: CellRangeAddress, is_shift_left: bool) -> bool:
        """
        Deletes cells in a spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cr_addr (CellRangeAddress): Cell range Address.
            is_shift_left (bool): If ``True`` then cell are shifted left; Otherwise, cells are shifted up.

        Returns:
            bool: ``True`` if cells are deleted; Otherwise, ``False``.
        """
        ...

    @overload
    @classmethod
    def delete_cells(cls, sheet: XSpreadsheet, range_name: str, is_shift_left: bool) -> bool:
        """
        Deletes cells in a spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_range (XCellRange): Cell range to delete.
            range_name (str): Range Name such as ``A1:D5``.
            is_shift_left (bool): If ``True`` then cell are shifted left; Otherwise, cells are shifted up.

        Returns:
            bool: ``True`` if cells are deleted; Otherwise, ``False``.
        """
        ...

    @overload
    @classmethod
    def delete_cells(cls, sheet: XSpreadsheet, range_obj: mRngObj.RangeObj, is_shift_left: bool) -> bool:
        """
        Deletes cells in a spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_obj (RangeObj): Range Object.
            is_shift_left (bool): If ``True`` then cell are shifted left; Otherwise, cells are shifted up.

        Returns:
            bool: ``True`` if cells are deleted; Otherwise, ``False``.
        """
        ...

    @overload
    @classmethod
    def delete_cells(
        cls, sheet: XSpreadsheet, col_start: int, row_start: int, col_end: int, row_end: int, is_shift_left: bool
    ) -> bool:
        """
        Deletes cells in a spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.
            is_shift_left (bool): If ``True`` then cell are shifted left; Otherwise, cells are shifted up.

        Returns:
            bool: ``True`` if cells are deleted; Otherwise, ``False``.
        """
        ...

    @classmethod
    def delete_cells(cls, *args, **kwargs) -> bool:
        """
        Deletes cells in a spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_range (XCellRange): Cell range to delete.
            cr_addr (CellRangeAddress): Cell range Address.
            range_name (str): Range Name such as ``A1:D5``.
            range_obj (RangeObj): Range Object.
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.
            is_shift_left (bool): If ``True`` then cell are shifted left; Otherwise, cells are shifted up.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_DELETING` :eventref:`src-docs-cell-event-deleting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_DELETED` :eventref:`src-docs-cell-event-deleted`

        Returns:
            bool: ``True`` if cells are deleted; Otherwise, ``False``.

        Note:
            Events args for this method have a ``cell`` type of ``XCellRange``

        Note:
            Event args ``event_data`` is a dictionary containing ``is_shift_left``.
        """
        kw = kwargs.copy()
        sheet = kw.get("sheet", None)
        lst_args = list(args)
        if sheet is None:
            sheet = lst_args[0]

        is_shift_left = kw.get("is_shift_left", None)
        # is_shift_left needs to be removed to pass args to get_cell_range()
        if is_shift_left is None:
            is_shift_left = bool(lst_args.pop())
        else:
            del kw["is_shift_left"]

        cell_range = None
        arg2 = kwargs.get("cell_range", None)
        if arg2 is None and len(lst_args) > 1:
            arg2 = lst_args[1]

        if arg2:
            cell_range = mLo.Lo.qi(XCellRange, arg2)

        if cell_range:
            return cls._delete_cells(sheet=sheet, cell_range=cell_range, is_shift_left=is_shift_left)

        cell_range = cls.get_cell_range(*lst_args, **kw)
        return cls._delete_cells(sheet=sheet, cell_range=cell_range, is_shift_left=is_shift_left)

    # endregion delete_cells()

    # region    clear_cells()
    @overload
    @classmethod
    def clear_cells(cls, sheet: XSpreadsheet, cell_range: XCellRange) -> bool:
        """
        Clears the specified contents of the cell range.

        If cell range of types ``VALUE``, ``DATETIME`` and ``STRING`` are cleared.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_range (XCellRange): Cell range.

        Returns:
            bool: ``True`` if cells are cleared; Otherwise, ``False``.
        """
        ...

    @overload
    @classmethod
    def clear_cells(cls, sheet: XSpreadsheet, cell_range: XCellRange, cell_flags: CellFlagsEnum) -> bool:
        """
        Clears the specified contents of the cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_range (XCellRange): Cell range.
            cell_flags (CellFlagsEnum): Flags that determine what to clear.

        Returns:
            bool: ``True`` if cells are cleared; Otherwise, ``False``.
        """
        ...

    @overload
    @classmethod
    def clear_cells(cls, sheet: XSpreadsheet, range_name: str) -> bool:
        """
        Clears the specified contents of the cell range.

        If cell range of types ``VALUE``, ``DATETIME`` and ``STRING`` are cleared.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range name such as ``A1:G3``.

        Returns:
            bool: ``True`` if cells are cleared; Otherwise, ``False``.
        """
        ...

    @overload
    @classmethod
    def clear_cells(cls, sheet: XSpreadsheet, range_name: str, cell_flags: CellFlagsEnum) -> bool:
        """
        Clears the specified contents of the cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range name such as ``A1:G3``.
            cell_flags (CellFlagsEnum): Flags that determine what to clear.

        Returns:
            bool: ``True`` if cells are cleared; Otherwise, ``False``.
        """
        ...

    @overload
    @classmethod
    def clear_cells(cls, sheet: XSpreadsheet, range_val: mRngObj.RangeObj) -> bool:
        """
        Clears the specified contents of the cell range.

        If cell range of types ``VALUE``, ``DATETIME`` and ``STRING`` are cleared.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_val (RangeObj): Range object.

        Returns:
            bool: ``True`` if cells are cleared; Otherwise, ``False``.
        """
        ...

    @overload
    @classmethod
    def clear_cells(cls, sheet: XSpreadsheet, range_val: mRngObj.RangeObj, cell_flags: CellFlagsEnum) -> bool:
        """
        Clears the specified contents of the cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_val (RangeObj): Range object.
            cell_flags (CellFlagsEnum): Flags that determine what to clear.

        Returns:
            bool: ``True`` if cells are cleared; Otherwise, ``False``.
        """
        ...

    @overload
    @classmethod
    def clear_cells(cls, sheet: XSpreadsheet, cr_addr: CellRangeAddress) -> bool:
        """
        Clears the specified contents of the cell range.

        If cell range of types ``VALUE``, ``DATETIME`` and ``STRING`` are cleared.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cr_addr (CellRangeAddress): Cell Range Address.

        Raises:
            MissingInterfaceError: If XSheetOperation interface cannot be obtained.
        Returns:
            bool: ``True`` if cells are cleared; Otherwise, ``False``.
        """
        ...

    @overload
    @classmethod
    def clear_cells(cls, sheet: XSpreadsheet, cr_addr: CellRangeAddress, cell_flags: CellFlagsEnum) -> bool:
        """
        Clears the specified contents of the cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cr_addr (CellRangeAddress): Cell Range Address.
            cell_flags (CellFlagsEnum): Flags that determine what to clear.

        Raises:
            MissingInterfaceError: If XSheetOperation interface cannot be obtained.
        Returns:
            bool: ``True`` if cells are cleared; Otherwise, ``False``.
        """
        ...

    @classmethod
    def clear_cells(cls, *args, **kwargs) -> bool:
        """
        Clears the specified contents of the cell range.

        If ``cell_flags`` is not specified then
        cell range of types ``VALUE``, ``DATETIME`` and ``STRING`` are cleared.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_range (XCellRange): Cell range.
            range_name (str): Range name such as ``A1:G3``.
            range_val (RangeObj): Range object.
            cr_addr (CellRangeAddress): Cell Range Address.
            cell_flags (CellFlagsEnum): Flags that determine what to clear.

        Raises:
            MissingInterfaceError: If XSheetOperation interface cannot be obtained.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_CLEARING` :eventref:`src-docs-cell-event-clearing`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_CLEARED` :eventref:`src-docs-cell-event-cleared`

        Returns:
            bool: ``True`` if cells are cleared; Otherwise, ``False``.

        Note:
            Events arg for this method have a ``cell`` type of ``XCellRange``.

            Events arg ``event_data`` is a dictionary containing ``cell_flags``.
        """
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("sheet", "cell_range", "range_name", "range_val", "cr_addr", "cell_flags")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("clear_cells() got an unexpected keyword argument")
            ka[1] = kwargs.get("sheet", None)
            keys = ("cell_range", "range_name", "range_val", "cr_addr")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("cell_flags", None)
            return ka

        if count not in (2, 3):
            raise TypeError("clear_cells() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 2:
            flags = CellFlagsEnum.VALUE | CellFlagsEnum.DATETIME | CellFlagsEnum.STRING
        else:
            if isinstance(kargs[3], int):
                flags = CellFlagsEnum(kargs[3])
            else:
                flags = cast(CellFlagsEnum, kargs[3])
        sht = cast(XSpreadsheet, kargs[1])
        rng_value = kargs[2]
        if isinstance(rng_value, (str, mRngObj.RangeObj)):
            rng = Calc.get_cell_range(sheet=sht, range_name=str(rng_value))
        elif mLo.Lo.is_uno_interfaces(rng_value, XCellRange):
            rng = rng_value
        else:
            rng = Calc.get_cell_range(sheet=sht, cr_addr=rng_value)

        cargs = CellCancelArgs(Calc.clear_cells.__qualname__)
        cargs.cells = rng
        cargs.sheet = sht
        cargs.event_data = {"cell_flags": flags}
        _Events().trigger(CalcNamedEvent.CELLS_CLEARING, cargs)
        if cargs.cancel:
            return False
        flags = cargs.event_data["cell_flags"]
        sheet_op = mLo.Lo.qi(XSheetOperation, cargs.cells, True)
        sheet_op.clearContents(flags.value)
        _Events().trigger(CalcNamedEvent.CELLS_CLEARED, CellArgs.from_args(cargs))
        return True

    # endregion clear_cells()

    # endregion ------------ insert/remove rows, columns, cells -----

    # region --------------- set/get values in cells -------------------
    # region    set_val()
    @staticmethod
    def _set_val_by_cell(value: Any, cell: XCell, styles: Sequence[StyleT] | None = None) -> None:
        """LO Safe Method"""
        if isinstance(value, (float, int)):
            cell.setValue(float(value))
        elif isinstance(value, str):
            cell.setFormula(str(value))
        else:
            mLo.Lo.print(f"Value is not a number or string: {value}")
        if styles is None:
            return
        supported_styles = (
            "com.sun.star.style.CharacterProperties",
            "com.sun.star.style.ParagraphProperties",
            "com.sun.star.table.CellProperties",
            "com.sun.star.sheet.SheetCellRange",
            "com.sun.star.sheet.SheetCell",
        )
        for style in styles:
            if style.support_service(*supported_styles):
                style.apply(cell)

    @classmethod
    def _set_val_by_cell_name(
        cls, value: object, sheet: XSpreadsheet, cell_name: str, styles: Sequence[StyleT] | None = None
    ) -> None:
        """LO Safe Method"""
        pos = cls.get_cell_position(cell_name)
        cls._set_val_by_col_row(value=value, sheet=sheet, col=pos.X, row=pos.Y, styles=styles)

    @classmethod
    def _set_val_by_col_row(
        cls, value: object, sheet: XSpreadsheet, col: int, row: int, styles: Sequence[StyleT] | None = None
    ) -> None:
        """LO Safe Method"""
        cell = cls.get_cell(sheet=sheet, col=col, row=row)
        cls._set_val_by_cell(value=value, cell=cell, styles=styles)

    @overload
    @classmethod
    def set_val(cls, value: object, cell: XCell) -> None:
        """
        Sets the value of a cell.

        |lo_safe|

        Args:
            value (object): Value for cell.
            cell (XCell): Cell to assign value.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_val(cls, value: object, cell: XCell, *, styles: Sequence[StyleT]) -> None:
        """
        Sets the value of a cell.

        |lo_safe|

        Args:
            value (object): Value for cell.
            cell (XCell): Cell to assign value.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_val(cls, value: object, sheet: XSpreadsheet, cell_name: str) -> None:
        """
        Sets the value of a cell.

        |lo_safe|

        Args:
            value (object): Value for cell.
            sheet (XSpreadsheet): Spreadsheet.
            cell_name (str): Name of cell to set value of such as ``B4``.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_val(cls, value: object, sheet: XSpreadsheet, cell_name: str, *, styles: Sequence[StyleT]) -> None:
        """
        Sets the value of a cell.

        |lo_safe|

        Args:
            value (object): Value for cell.
            sheet (XSpreadsheet): Spreadsheet.
            cell_name (str): Name of cell to set value of such as ``B4``.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_val(cls, value: object, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj) -> None:
        """
        Sets the value of a cell.

        |lo_safe|

        Args:
            value (object): Value for cell.
            sheet (XSpreadsheet): Spreadsheet.
            cell_obj (CellObj): Cell Object.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_val(
        cls, value: object, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj, *, styles: Sequence[StyleT]
    ) -> None:
        """
        Sets the value of a cell.

        |lo_safe|

        Args:
            value (object): Value for cell.
            sheet (XSpreadsheet): Spreadsheet.
            cell_obj (CellObj): Cell Object.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_val(cls, value: object, sheet: XSpreadsheet, col: int, row: int) -> None:
        """
        Sets the value of a cell.

        |lo_safe|

        Args:
            value (object): Value for cell.
            sheet (XSpreadsheet): Spreadsheet.
            col (int): Cell column as zero-based integer.
            row (int): Cell row as zero-based integer.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_val(cls, value: object, sheet: XSpreadsheet, col: int, row: int, *, styles: Sequence[StyleT]) -> None:
        """
        Sets the value of a cell.

        |lo_safe|

        Args:
            value (object): Value for cell.
            sheet (XSpreadsheet): Spreadsheet.
            col (int): Cell column as zero-based integer.
            row (int): Cell row as zero-based integer.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell.

        Returns:
            None:
        """
        ...

    @classmethod
    def set_val(cls, *args, **kwargs) -> None:
        """
        Sets the value of a cell.

        |lo_safe|

        Args:
            value (object): Value for cell.
            cell (XCell): Cell to assign value.
            sheet (XSpreadsheet): Spreadsheet.
            cell_name (str): Name of cell to set value of such as ``B4``.
            cell_obj (CellObj): Cell Object.
            col (int): Cell column as zero-based integer.
            row (int): Cell row as zero-based integer.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_style_cell`
            - :ref:`help_calc_format_direct_cell`

        .. versionchanged:: 0.9.1
            Added overloads for styles.
        """
        styles = kwargs.pop("styles", None)
        ordered_keys = (1, 2, 3, 4)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("value", "cell", "sheet", "cell_name", "cell_obj", "col", "row")
            check = all(key in valid_keys for key in kwargs)
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

            keys = ("cell_name", "cell_obj", "col")
            for key in keys:
                if key in kwargs:
                    ka[3] = kwargs[key]
                    break
            if count == 3:
                return ka
            ka[4] = kwargs.get("row")
            return ka

        if count not in (2, 3, 4):
            raise TypeError("set_val() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 2:
            cls._set_val_by_cell(value=kargs[1], cell=kargs[2], styles=styles)
        elif count == 3:
            cls._set_val_by_cell_name(value=kargs[1], sheet=kargs[2], cell_name=str(kargs[3]), styles=styles)
        elif count == 4:
            cls._set_val_by_col_row(value=kargs[1], sheet=kargs[2], col=kargs[3], row=kargs[4], styles=styles)

    # endregion    set_val()

    @staticmethod
    def convert_to_float(val: Any) -> float:
        """
        Converts value to float.

        |lo_safe|

        Args:
            val (Any): Value to convert

        Returns:
            float: value converted to float. 0.0 is returned if conversion fails.
        """
        if val is None:
            mLo.Lo.print("Value is null; using 0")
            return 0.0
        try:
            return float(val)
        except ValueError:
            mLo.Lo.print(f"Could not convert {val} to double; using 0")
            return 0.0

    convert_to_double = convert_to_float

    @staticmethod
    def get_type_enum(cell: XCell) -> Calc.CellTypeEnum:
        """
        Gets enum representing the Type.

        |lo_safe|

        Args:
            cell (XCell): Cell to get type of

        Returns:
            CellTypeEnum: Enum of cell type
        """
        t = cell.getType()
        if t == CellContentType.EMPTY:
            return Calc.CellTypeEnum.EMPTY
        if t == CellContentType.VALUE:
            return Calc.CellTypeEnum.VALUE
        if t == CellContentType.TEXT:
            return Calc.CellTypeEnum.TEXT
        if t == CellContentType.FORMULA:
            return Calc.CellTypeEnum.FORMULA
        mLo.Lo.print("Unknown cell type")
        return Calc.CellTypeEnum.UNKNOWN

    @classmethod
    def get_type_string(cls, cell: XCell) -> str:
        """
        Gets String representing the Type.

        |lo_safe|

        Args:
            cell (XCell): Cell to get type of

        Returns:
            str: String of cell type
        """
        t = cls.get_type_enum(cell=cell)
        return str(t)

    # region    get_val()

    @classmethod
    def _get_val_by_cell(cls, cell: XCell) -> object | None:
        """LO Safe Method."""
        t = cell.getType()
        if t == CellContentType.EMPTY:
            return None
        if t == CellContentType.VALUE:
            return cls.convert_to_float(cell.getValue())
        if t in (CellContentType.TEXT, CellContentType.FORMULA):
            return cell.getFormula()
        mLo.Lo.print("Unknown cell type; returning None")
        return None

    @classmethod
    def _get_val_by_col_row(cls, sheet: XSpreadsheet, col: int, row: int) -> object | None:
        """LO Safe Method"""
        xcell = cls.get_cell(sheet=sheet, col=col, row=row)
        return cls._get_val_by_cell(cell=xcell)

    @classmethod
    def _get_val_by_cell_name(cls, sheet: XSpreadsheet, cell_name: str) -> object | None:
        """Lo Safe Method"""
        pos = cls.get_cell_position(cell_name)
        return cls._get_val_by_col_row(sheet=sheet, col=pos.X, row=pos.Y)

    @classmethod
    def _get_val_by_cell_addr(cls, sheet: XSpreadsheet, addr: CellAddress) -> object | None:
        """LO Safe Method."""
        if addr is None:
            return None
        return cls._get_val_by_col_row(sheet=sheet, col=addr.Column, row=addr.Row)

    @overload
    @classmethod
    def get_val(cls, cell: XCell) -> Any:
        """
        Gets cell value.

        |lo_safe|

        Args:
            cell (XCell): cell to get value of.

        Returns:
            Any: Cell value cell has a value; Otherwise, ``None``.
        """
        ...

    @overload
    @classmethod
    def get_val(cls, sheet: XSpreadsheet, addr: CellAddress) -> Any:
        """
        Gets cell value.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            addr (CellAddress): Address of cell.

        Returns:
            Any: Cell value cell has a value; Otherwise, ``None``.
        """
        ...

    @overload
    @classmethod
    def get_val(cls, sheet: XSpreadsheet, cell_name: str) -> Any:
        """
        Gets cell value.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_name (str): Name of cell such as ``B4``.

        Returns:
            Any: Cell value cell has a value; Otherwise, ``None``.
        """
        ...

    @overload
    @classmethod
    def get_val(cls, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj) -> Any:
        """
        Gets cell value.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_obj (CellObj): Cell Object.

        Returns:
            Any: Cell value cell has a value; Otherwise, ``None``.
        """
        ...

    @overload
    @classmethod
    def get_val(cls, sheet: XSpreadsheet, col: int, row: int) -> Any:
        """
        Gets cell value.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            col (int): Cell zero-based column.
            row (int): Cell zero-base row.

        Returns:
            Any: Cell value cell has a value; Otherwise, ``None``.
        """
        ...

    @classmethod
    def get_val(cls, *args, **kwargs) -> Any:
        """
        Gets cell value.

        |lo_safe|

        Args:
            cell (XCell): cell to get value of.
            sheet (XSpreadsheet): Spreadsheet.
            addr (CellAddress): Address of cell.
            cell_name (str): Name of cell such as ``B4``.
            cell_obj (CellObj): Cell Object.
            col (int): Cell zero-based column.
            row (int): Cell zero-base row.

        Returns:
            Any: Cell value cell has a value; Otherwise, ``None``.
        """
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("sheet", "cell", "cell_name", "cell_obj", "addr", "col", "row")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_val() got an unexpected keyword argument")
            keys = ("sheet", "cell")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            keys = ("addr", "cell_name", "cell_obj", "col")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("row", None)
            return ka

        if count not in (1, 2, 3):
            raise TypeError("get_val() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        first_arg = mLo.Lo.qi(XSpreadsheet, kargs[1])
        if first_arg is None:
            # can only be: get_val(cell: XCell)
            return None if count != 1 else cls._get_val_by_cell(cell=kargs[1])
        if count == 2:
            if mInfo.Info.is_instance(kargs[2], (str, mCellObj.CellObj)):
                #   get_val(sheet: XSpreadsheet, cell_name: str)
                return cls._get_val_by_cell_name(sheet=kargs[1], cell_name=str(kargs[2]))

            #   get_val(sheet: XSpreadsheet, addr: CellAddress)
            return cls._get_val_by_cell_addr(sheet=kargs[1], addr=kargs[2])

        if count == 3:
            #   get_val(sheet: XSpreadsheet, col: int, row: int)
            return cls._get_val_by_col_row(sheet=kargs[1], col=kargs[2], row=kargs[3])
        return None

    # endregion get_val()

    # region    get_num()

    # cell: XCell
    @overload
    @classmethod
    def get_num(cls, cell: XCell) -> float:
        """
        Get cell value a float.

        |lo_safe|

        Args:
            cell (XCell): Cell to get value of.

        Returns:
            float: Cell value as float. If cell value cannot be converted then ``0.0`` is returned.
        """
        ...

    @overload
    @classmethod
    def get_num(cls, sheet: XSpreadsheet, cell_name: str) -> float:
        """
        Get cell value a float.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_name (str): Cell name such as ``B4``.

        Returns:
            float: Cell value as float. If cell value cannot be converted then ``0.0`` is returned.
        """
        ...

    @overload
    @classmethod
    def get_num(cls, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj) -> float:
        """
        Get cell value a float.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_obj (CellObj): Cell Object.

        Returns:
            float: Cell value as float. If cell value cannot be converted then ``0.0`` is returned.
        """
        ...

    @overload
    @classmethod
    def get_num(cls, sheet: XSpreadsheet, addr: CellAddress) -> float:
        """
        Get cell value a float.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            addr (CellAddress): Cell Address.

        Returns:
            float: Cell value as float. If cell value cannot be converted then ``0.0`` is returned.
        """
        ...

    @overload
    @classmethod
    def get_num(cls, sheet: XSpreadsheet, col: int, row: int) -> float:
        """
        Get cell value a float.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            col (int): Cell zero-base column number.
            row (int): Cell zero-base row number.

        Returns:
            float: Cell value as float. If cell value cannot be converted then ``0.0`` is returned.
        """
        ...

    @classmethod
    def get_num(cls, *args, **kwargs) -> float:
        """
        Get cell value a float.

        |lo_safe|

        Args:
            cell (XCell): Cell to get value of.
            sheet (XSpreadsheet): Spreadsheet.
            cell_name (str): Cell name such as ``B4``.
            cell_obj (CellObj): Cell Object.
            addr (CellAddress): Cell Address.
            col (int): Cell zero-base column number.
            row (int): Cell zero-base row number.

        Returns:
            float: Cell value as float. If cell value cannot be converted then ``0.0`` is returned.
        """
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("sheet", "cell", "cell_name", "cell_obj", "addr", "col", "row")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_num() got an unexpected keyword argument")
            keys = ("sheet", "cell")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            keys = ("cell_name", "cell_obj", "addr", "col")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("row", None)
            return ka

        if count not in (1, 2, 3):
            raise TypeError("get_num() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            return cls.convert_to_float(cls.get_val(cell=kargs[1]))

        if count == 3:
            return cls.convert_to_float(cls.get_val(sheet=kargs[1], col=kargs[2], row=kargs[3]))
        if count == 2:
            return cls.convert_to_float(cls.get_val(kargs[1], kargs[2]))
        return 0.0

    # endregion get_num()

    # region    get_string()
    @overload
    @classmethod
    def get_string(cls, cell: XCell) -> str:
        """
        Gets the value of a cell as a string.

        |lo_safe|

        Args:
            cell (XCell): Cell to get value of.

        Returns:
            str: Cell value as string.
        """
        ...

    @overload
    @classmethod
    def get_string(cls, sheet: XSpreadsheet, cell_name: str) -> str:
        """
        Gets the value of a cell as a string.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_name (str): Name of cell to get the value of such as ``B4``.

        Returns:
            str: Cell value as string.
        """
        ...

    @overload
    @classmethod
    def get_string(cls, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj) -> str:
        """
        Gets the value of a cell as a string.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_obj (CellObj): Cell Object.

        Returns:
            str: Cell value as string.
        """
        ...

    @overload
    @classmethod
    def get_string(cls, sheet: XSpreadsheet, addr: CellAddress) -> str:
        """
        Gets the value of a cell as a string.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            addr (CellAddress): Cell address.

        Returns:
            str: Cell value as string.
        """
        ...

    @overload
    @classmethod
    def get_string(cls, sheet: XSpreadsheet, col: int, row: int) -> str:
        """
        Gets the value of a cell as a string.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            col (int): Cell zero-based column number.
            row (int): Cell zero-based row number.

        Returns:
            str: Cell value as string.
        """
        ...

    @classmethod
    def get_string(cls, *args, **kwargs) -> str:
        """
        Gets the value of a cell as a string.

        |lo_safe|

        Args:
            cell (XCell): Cell to get value of.
            sheet (XSpreadsheet): Spreadsheet.
            cell_name (str): Name of cell to get the value of such as ``B4``.
            addr (CellAddress): Cell address.
            cell_obj (CellObj): Cell Object.
            col (int): Cell zero-based column number.
            row (int): Cell zero-based row number.

        Returns:
            str: Cell value as string.
        """
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("cell", "sheet", "cell_name", "cell_obj", "addr", "col", "row")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_string() got an unexpected keyword argument")
            keys = ("cell", "sheet")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            keys = ("cell_name", "cell_obj", "addr", "col")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("row", None)
            return ka

        if count not in (1, 2, 3):
            raise TypeError("get_string() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        def convert(obj) -> str:
            return "" if obj is None else str(obj)

        if count == 1:
            return convert(cls.get_val(cell=kargs[1]))

        if count == 3:
            return convert(cls.get_val(sheet=kargs[1], col=kargs[2], row=kargs[3]))
        if count == 2:
            if isinstance(kargs[2], (str, mCellObj.CellObj)):
                return convert(cls.get_val(sheet=kargs[1], cell_name=str(kargs[2])))
            return convert(cls.get_val(sheet=kargs[1], addr=kargs[2]))
        return ""

    # endregion get_string()

    # endregion ------------ set/get values in cells -----------------

    # region --------------- set/get values in 2D array ----------------

    # region    set_array()
    @classmethod
    def _set_array_doc_addr(
        cls, values: Table, doc: XSpreadsheetDocument, addr: CellAddress, styles: Sequence[StyleT] | None = None
    ) -> None:
        """Lo Safe Method."""
        v_len = len(values)
        if v_len == 0:
            mLo.Lo.print("Values has not data")
            return
        sheet = cls._get_sheet_index(doc=doc, index=addr.Sheet)
        col_end = addr.Column + (len(values[0]) - 1)
        row_end = addr.Row + (v_len - 1)
        cell_range = cls._get_cell_range_col_row(
            sheet=sheet, start_col=addr.Column, start_row=addr.Row, end_col=col_end, end_row=row_end
        )
        if styles is None:
            cls.set_cell_range_array(cell_range=cell_range, values=values)
        else:
            cls.set_cell_range_array(cell_range=cell_range, values=values, styles=styles)

    @overload
    @classmethod
    def set_array(cls, values: Table, cell_range: XCellRange) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            cell_range (XCellRange): Range in spreadsheet to insert data.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_array(cls, values: Table, cell_range: XCellRange, *, styles: Sequence[StyleT]) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            cell_range (XCellRange): Range in spreadsheet to insert data.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_array(cls, values: Table, sheet: XSpreadsheet, name: str) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            sheet (XSpreadsheet): Spreadsheet.
            name (str): Range name such as 'A1:D4' or cell name such as ``B4``.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_array(cls, values: Table, sheet: XSpreadsheet, name: str, *, styles: Sequence[StyleT]) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            sheet (XSpreadsheet): Spreadsheet.
            name (str): Range name such as 'A1:D4' or cell name such as ``B4``.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_array(cls, values: Table, sheet: XSpreadsheet, range_obj: mRngObj.RangeObj) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            sheet (XSpreadsheet): Spreadsheet.
            range_obj (RangeObj): Range Object.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_array(
        cls, values: Table, sheet: XSpreadsheet, range_obj: mRngObj.RangeObj, *, styles: Sequence[StyleT]
    ) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            sheet (XSpreadsheet): Spreadsheet.
            range_obj (RangeObj): Range Object.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_array(cls, values: Table, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            sheet (XSpreadsheet): Spreadsheet.
            cell_obj (CellObj): Cell Object.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_array(
        cls, values: Table, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj, *, styles: Sequence[StyleT]
    ) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            sheet (XSpreadsheet): Spreadsheet.
            cell_obj (CellObj): Cell Object.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_array(cls, values: Table, doc: XSpreadsheetDocument, addr: CellAddress) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            doc (XSpreadsheetDocument): Spreadsheet Document.
            addr (CellAddress): Address to insert data.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_array(
        cls, values: Table, doc: XSpreadsheetDocument, addr: CellAddress, *, styles: Sequence[StyleT]
    ) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            doc (XSpreadsheetDocument): Spreadsheet Document.
            addr (CellAddress): Address to insert data.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_array(
        cls,
        values: Table,
        sheet: XSpreadsheet,
        col_start: int,
        row_start: int,
        col_end: int,
        row_end: int,
    ) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            sheet (XSpreadsheet): Spreadsheet.
            col_start (int): Zero-base Start Column.
            row_start (int): Zero-base Start Row.
            col_end (int): Zero-base End Column.
            row_end (int): Zero-base End Row.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_array(
        cls,
        values: Table,
        sheet: XSpreadsheet,
        col_start: int,
        row_start: int,
        col_end: int,
        row_end: int,
        *,
        styles: Sequence[StyleT],
    ) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            sheet (XSpreadsheet): Spreadsheet.
            col_start (int): Zero-base Start Column.
            row_start (int): Zero-base Start Row.
            col_end (int): Zero-base End Column.
            row_end (int): Zero-base End Row.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @classmethod
    def set_array(cls, *args, **kwargs) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            cell_range (XCellRange): Range in spreadsheet to insert data.
            sheet (XSpreadsheet): Spreadsheet.
            name (str): Range name such as 'A1:D4' or cell name such as ``B4``.
            range_obj (RangeObj): Range Object.
            cell_obj (CellObj): Cell Object.
            doc (XSpreadsheetDocument): Spreadsheet Document.
            addr (CellAddress): Address to insert data.
            col_start (int): Zero-base Start Column.
            row_start (int): Zero-base Start Row.
            col_end (int): Zero-base End Column.
            row_end (int): Zero-base End Row.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_style_cell`
            - :ref:`help_calc_format_direct_cell`

        .. versionchanged:: 0.9.1
            Added overloads for styles.
        """
        styles = cast(Sequence[StyleT], kwargs.pop("styles", ()))
        ordered_keys = (1, 2, 3, 4, 5, 6)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = (
                "values",
                "cell_range",
                "sheet",
                "doc",
                "name",
                "range_obj",
                "cell_obj",
                "col_start",
                "addr",
                "row_start",
                "col_end",
                "row_end",
            )
            check = all(key in valid_keys for key in kwargs)
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
            keys = ("name", "col_start", "addr", "range_obj", "cell_obj")
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

        if count not in (2, 3, 6):
            raise TypeError("set_array() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 2:
            #  set_array(values: Sequence[Sequence[object]], cell_range: XCellRange)
            cls.set_cell_range_array(cell_range=kargs[2], values=kargs[1], styles=styles)
            return
        if count == 3:
            arg1 = kargs[1]
            arg2 = kargs[2]
            arg3 = kargs[3]

            if isinstance(arg3, str):
                # set_array(values: Sequence[Sequence[object]], sheet: XSpreadsheet, name: str)
                if cls.is_cell_range_name(arg3):
                    cls._set_array_range(sheet=arg2, range_name=cls.get_safe_rng_str(arg3), values=arg1, styles=styles)
                    return
                else:
                    cls._set_array_cell(sheet=arg2, cell_name=arg3, values=arg1, styles=styles)
                    return
            elif isinstance(arg3, mRngObj.RangeObj):
                cls._set_array_range(sheet=arg2, range_name=arg3, values=arg1, styles=styles)
            elif isinstance(arg3, mCellObj.CellObj):
                cls._set_array_cell(sheet=arg2, cell_name=arg3, values=arg1, styles=styles)
            else:
                cls._set_array_doc_addr(values=arg1, doc=arg2, addr=arg3, styles=styles)
            return
        if count == 6:
            #  def set_array(values: Sequence[Sequence[object]], sheet: XSpreadsheet, col_start: int, row_start: int, col_end:int, row_end: int)
            cell_range = cls._get_cell_range_col_row(
                sheet=kargs[2], start_col=kargs[3], start_row=kargs[4], end_col=kargs[5], end_row=kargs[6]
            )
            cls.set_cell_range_array(cell_range=cell_range, values=kargs[1], styles=styles)
        return

    # endregion set_array()

    # region set_array_range()

    @classmethod
    def _set_array_range(
        cls,
        sheet: XSpreadsheet,
        range_name: str | mRngObj.RangeObj,
        values: Table,
        styles: Sequence[StyleT] | None = None,
    ) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet
            range_name (str): Range to insert data such as 'A1:E12'
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        v_len = len(values)
        if v_len == 0:
            mLo.Lo.print("Values has not data")
            return
        cell_range = cls.get_cell_range(sheet, range_name)
        if styles is None:
            cls.set_cell_range_array(cell_range=cell_range, values=values)
        else:
            cls.set_cell_range_array(cell_range=cell_range, values=values, styles=styles)

    @overload
    @classmethod
    def set_array_range(cls, sheet: XSpreadsheet, range_name: str, values: Table) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range to insert data such as ``A1:E12``.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_array_range(cls, sheet: XSpreadsheet, range_name: str, values: Table, *, styles: Sequence[StyleT]) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range to insert data such as ``A1:E12``.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_array_range(cls, sheet: XSpreadsheet, range_obj: mRngObj.RangeObj, values: Table) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_obj (RangeObj): Range Object.
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_array_range(
        cls, sheet: XSpreadsheet, range_obj: mRngObj.RangeObj, values: Table, *, styles: Sequence[StyleT]
    ) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_obj (RangeObj): Range Object.
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @classmethod
    def set_array_range(cls, *args, **kwargs) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range to insert data such as ``A1:E12``.
            range_obj (RangeObj): Range Object.
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:

        .. versionchanged:: 0.9.1
            Added overloads for styles.
        """
        styles = kwargs.pop("styles", None)
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("sheet", "range_name", "range_obj", "values")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("set_array_range() got an unexpected keyword argument")
            ka[1] = kwargs.get("sheet", None)
            keys = ("range_name", "range_obj")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            ka[3] = kwargs.get("values", None)
            return ka

        if count != 3:
            raise TypeError("set_array_range() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        cls._set_array_range(sheet=kargs[1], range_name=str(kargs[2]), values=kargs[3], styles=styles)

    # endregion set_array_range()

    # region set_cell_range_array()
    @overload
    @staticmethod
    def set_cell_range_array(cell_range: XCellRange, values: Table) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.

        Returns:
            None:
        """
        ...

    @overload
    @staticmethod
    def set_cell_range_array(cell_range: XCellRange, values: Table, styles: Sequence[StyleT]) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @staticmethod
    def set_cell_range_array(cell_range: XCellRange, values: Table, styles: Sequence[StyleT] | None = None) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_style_cell`
            - :ref:`help_calc_format_direct_cell`

        .. versionchanged:: 0.9.1
            Added overloads for styles.
        """

        v_len = len(values)
        if v_len == 0:
            mLo.Lo.print("Values has not data")
            return
        cr_data = mLo.Lo.qi(XCellRangeData, cell_range)
        if cr_data is None:
            return
        cr_data.setDataArray(values)  # type: ignore
        if styles is None:
            return
        supported_styles = (
            "com.sun.star.style.CharacterProperties",
            "com.sun.star.style.ParagraphProperties",
            "com.sun.star.table.CellProperties",
            "com.sun.star.sheet.SheetCellRange",
            "com.sun.star.sheet.SheetCell",
        )

        for style in styles:
            if style.support_service(*supported_styles):
                style.apply(cell_range)

    # endregion set_cell_range_array()

    # region set_array_cell()

    @classmethod
    def _set_array_cell(
        cls,
        sheet: XSpreadsheet,
        cell_name: str | mCellObj.CellObj,
        values: Table,
        styles: Sequence[StyleT] | None = None,
    ) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cell_name (str): Cell Name such as 'A1'
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.
        """
        v_len = len(values)
        if v_len == 0:
            mLo.Lo.print("Values has not data")
            return
        pos = cls.get_cell_position(cell_name)
        col_end = pos.X + (len(values[0]) - 1)
        row_end = pos.Y + (v_len - 1)
        cell_range = cls._get_cell_range_col_row(
            sheet=sheet, start_col=pos.X, start_row=pos.Y, end_col=col_end, end_row=row_end
        )
        if styles is None:
            cls.set_cell_range_array(cell_range=cell_range, values=values)
        else:
            cls.set_cell_range_array(cell_range=cell_range, values=values, styles=styles)

    @overload
    @classmethod
    def set_array_cell(cls, sheet: XSpreadsheet, range_name: str, values: Table) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range to insert data such as ``A1:E12``.
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_array_cell(cls, sheet: XSpreadsheet, range_name: str, values: Table, *, styles: Sequence[StyleT]) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range to insert data such as ``A1:E12``.
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_array_cell(cls, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj, values: Table) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_obj (CellObj): Range Object.
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_array_cell(
        cls, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj, values: Table, *, styles: Sequence[StyleT]
    ) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_obj (CellObj): Range Object.
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @classmethod
    def set_array_cell(cls, *args, **kwargs) -> None:
        """
        Inserts array of data into spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range to insert data such as ``A1:E12``.
            cell_obj (CellObj): Range Object,
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_style_cell`
            - :ref:`help_calc_format_direct_cell`

        .. versionchanged:: 0.9.1
            Added overloads for styles.
        """
        styles = kwargs.pop("styles", None)
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("sheet", "range_name", "cell_obj", "values")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("set_array_range() got an unexpected keyword argument")
            ka[1] = kwargs.get("sheet", None)
            keys = ("range_name", "cell_obj")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            ka[3] = ka.get("values")
            return ka

        if count != 3:
            raise TypeError("set_array_range() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        cls._set_array_cell(sheet=kargs[1], cell_name=kargs[2], values=kargs[3], styles=styles)

    # endregion set_array_cell()

    # region get_array()

    @overload
    @classmethod
    def get_array(cls, cell_range: XCellRange) -> TupleArray:
        """
        Gets Array of data from a spreadsheet.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell range to get data from.

        Returns:
            TupleArray: Resulting data array.
        """
        ...

    @overload
    @classmethod
    def get_array(cls, sheet: XSpreadsheet, range_name: str) -> TupleArray:
        """
        Gets Array of data from a spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet
            range_name (str): Range of data to get such as "A1:E16"

        Returns:
            TupleArray: Resulting data array.
        """
        ...

    @overload
    @classmethod
    def get_array(cls, sheet: XSpreadsheet, range_obj: mRngObj.RangeObj) -> TupleArray:
        """
        Gets Array of data from a spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet
            range_obj (RangeObj): Range object

        Returns:
            TupleArray: Resulting data array.
        """
        ...

    @overload
    @classmethod
    def get_array(cls, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj) -> TupleArray:
        """
        Gets Array of data from a spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cell_obj (CellObj): Cell Object

        Returns:
            TupleArray: Resulting data array.
        """
        ...

    @classmethod
    def get_array(cls, *args, **kwargs) -> TupleArray:
        """
        Gets Array of data from a spreadsheet.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell range to get data from.
            sheet (XSpreadsheet): Spreadsheet
            range_name (str): Range of data to get such as "A1:E16"
            range_obj (RangeObj): Range object
            cell_obj (CellObj): Cell Object

        Raises:
            MissingInterfaceError: if interface is missing

        Returns:
            TupleArray: Resulting data array.
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("cell_range", "sheet", "range_name", "range_obj", "cell_obj")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_array() got an unexpected keyword argument")
            keys = ("cell_range", "sheet")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            keys = ("range_name", "range_obj", "cell_obj")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            return ka

        if count not in (1, 2):
            raise TypeError("get_array() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            cell_range = cast(XCellRange, kargs[1])
        else:
            cell_range = cls.get_cell_range(kargs[1], kargs[2])

        cr_data = mLo.Lo.qi(XCellRangeData, cell_range, raise_err=True)
        return cr_data.getDataArray()

    # endregion get_array()

    # region print_array()

    @overload
    @staticmethod
    def print_array(vals: Table) -> None: ...

    @overload
    @staticmethod
    def print_array(vals: Table, format_opt: FormatterTable) -> None: ...

    @staticmethod
    def print_array(vals: Table, format_opt: FormatterTable | None = None) -> None:
        """
        Prints a 2-Dimensional array to console.

        |lo_safe|

        Args:
            vals (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            format_opt (FormatterTable, optional): Optional format used to format values when printing to console such as ``FormatterTable(format=".2f")``

        Returns:
            None:

        See Also:
            - :ref:`ch21_format_data_console`
            - :py:data:`~.type_var.Table`

        .. versionchanged:: 0.6.6
            Added ``format_opt`` parameter

        .. versionchanged:: 0.6.10

            Removed cancel event args.
        """
        row_len = len(vals)
        if row_len == 0:
            print("No data in array to print")
            return
        col_len = len(vals[0])
        print(f"Row x Column size: {row_len} x {col_len}")

        if format_opt:
            for i, row in enumerate(vals):
                col_str = format_opt.get_formatted(idx_row=i, row_data=row)
                print(col_str)
        else:
            for row in vals:
                col_str = "  ".join([str(cell) for cell in row])
                print(col_str)
        print()

    # endregion print_array()

    # region get_float_array()

    @overload
    @classmethod
    def get_float_array(cls, cell_range: XCellRange) -> FloatTable:
        """
        Gets a 2-Dimensional List of floats.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell range to get data from.

        Returns:
            FloatTable: 2-Dimensional List of floats.
        """
        ...

    @overload
    @classmethod
    def get_float_array(cls, sheet: XSpreadsheet, range_name: str) -> FloatTable:
        """
        Gets a 2-Dimensional List of floats.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet to get the float values from.
            range_name (str): Range to get array of floats from such as ``A1:E18``.

        Returns:
            FloatTable: 2-Dimensional List of floats.
        """
        ...

    @overload
    @classmethod
    def get_float_array(cls, sheet: XSpreadsheet, range_obj: mRngObj.RangeObj) -> FloatTable:
        """
        Gets a 2-Dimensional List of floats.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet to get the float values from.
            range_obj (RangeObj): Range object.

        Returns:
            FloatTable: 2-Dimensional List of floats.
        """
        ...

    @overload
    @classmethod
    def get_float_array(cls, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj) -> FloatTable:
        """
        Gets a 2-Dimensional List of floats.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet to get the float values from.
            cell_obj (CellObj): Cell Object.

        Returns:
            FloatTable: 2-Dimensional List of floats.
        """
        ...

    @classmethod
    def get_float_array(cls, *args, **kwargs) -> FloatTable:
        """
        Gets a 2-Dimensional List of floats.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell range to get data from.
            sheet (XSpreadsheet): Spreadsheet to get the float values from.
            range_name (str): Range to get array of floats from such as ``A1:E18``.
            range_obj (RangeObj): Range object.
            cell_obj (CellObj): Cell Object.

        Returns:
            FloatTable: 2-Dimensional List of floats.
        """
        return cls._convert_to_floats_2d(cls.get_array(*args, **kwargs))

    # endregion get_float_array()

    get_doubles_array = get_float_array

    # region    convert_to_floats()

    @classmethod
    def _convert_to_floats_1d(cls, vals: Sequence[object]) -> FloatList:
        """LO Safe Method."""
        return [cls.convert_to_float(val) for val in vals]

    @classmethod
    def _convert_to_floats_2d(cls, vals: Sequence[Sequence[object]]) -> FloatTable:
        """LO Safe Method."""
        row_len = len(vals)
        if row_len == 0:
            return []
        col_len = len(vals[0])

        doubles = mTblHelper.TableHelper.make_2d_array(num_rows=row_len, num_cols=col_len)
        for row, col in itertools.product(range(row_len), range(col_len)):
            doubles[row][col] = cls.convert_to_float(vals[row][col])
        return doubles

    @overload
    @classmethod
    def convert_to_floats(cls, vals: Column) -> FloatList:
        """
        Converts a 1-Dimensional array into List of float.

        |lo_safe|

        Args:
            vals (Column): Sequence to convert to floats.

        Returns:
            FloatList: vals converted to float
        """
        ...

    @overload
    @classmethod
    def convert_to_floats(cls, vals: Row) -> FloatList:
        """
        Converts a 1-Dimensional array into List of float.

        |lo_safe|

        Args:
            vals (Row): Sequence to convert to floats.

        Returns:
            FloatList: vals converted to float
        """
        ...

    @overload
    @classmethod
    def convert_to_floats(cls, vals: Table) -> FloatTable:
        """
        Converts a 2-Dimensional array into List of float.

        |lo_safe|

        Args:
            vals (Table): 2-Dimensional list to convert to floats

        Returns:
            FloatTable: 2-Dimensional list of floats.
        """
        ...

    @classmethod
    def convert_to_floats(cls, vals: Row | Table) -> FloatList | FloatTable:
        """
        Converts a 1d or 2d array into List of float.

        |lo_safe|

        Args:
            vals (Row | Table): List or 2-Dimensional list to convert to floats.

        Returns:
            FloatList | FloatTable: vals converted to float
        """
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
    @classmethod
    def set_col(cls, sheet: XSpreadsheet, values: Column, cell_name: str) -> None:
        """
        Inserts a column of data into spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            values (Column): Column Data.
            cell_name (str): Name of Cell to begin the insert such as ``A1``.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_col(cls, sheet: XSpreadsheet, values: Column, cell_obj: mCellObj.CellObj) -> None:
        """
        Inserts a column of data into spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            values (Column): Column Data.
            cell_obj (CellObj): Cell Object.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_col(cls, sheet: XSpreadsheet, values: Column, col_start: int, row_start: int) -> None:
        """
        Inserts a column of data into spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            values (Column): Column Data.
            col_start (int): Zero-base column index.
            row_start (int): Zero-base row index.

        Returns:
            None:
        """
        ...

    @classmethod
    def set_col(cls, *args, **kwargs) -> None:
        """
        Inserts a column of data into spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            values (Column): Column Data.
            cell_name (str): Name of Cell to begin the insert such as ``A1``.
            cell_obj (CellObj): Cell Object.
            col_start (int): Zero-base column index.
            row_start (int): Zero-base row index.

        Returns:
            None:
        """
        ordered_keys = (1, 2, 3, 4)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("sheet", "values", "cell_name", "cell_obj", "col_start", "row_start")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("set_col() got an unexpected keyword argument")
            ka[1] = kwargs.get("sheet", None)
            ka[2] = kwargs.get("values", None)
            keys = ("cell_name", "cell_obj", "col_start")
            for key in keys:
                if key in kwargs:
                    ka[3] = kwargs[key]
                    break
            if count == 3:
                return ka
            ka[4] = kwargs.get("row_start", None)
            return ka

        if count not in (3, 4):
            raise TypeError("set_col() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 3:
            pos = cls.get_cell_position(str(kargs[3]))
            x = pos.X
            y = pos.Y
        else:
            x = kargs[3]
            y = kargs[4]
        values = cast(Sequence[Any], kargs[2])
        val_len = len(values)  # values

        cell_range = cls.get_cell_range(sheet=kargs[1], col_start=x, row_start=y, col_end=x, row_end=y + val_len - 1)
        for val in range(val_len):
            xcell = cls.get_cell(cell_range=cell_range, col=0, row=val)
            cls.set_val(cell=xcell, value=values[val])

    # endregion set_col()

    # region    set_row()
    @overload
    @classmethod
    def set_row(cls, sheet: XSpreadsheet, values: Row, cell_name: str) -> None:
        """
        Inserts a row of data into spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            values (Row): Row Data.
            cell_name (str): Name of Cell to begin the insert such as ``A1``.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_row(cls, sheet: XSpreadsheet, values: Row, cell_obj: mCellObj.CellObj) -> None:
        """
        Inserts a row of data into spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            values (Row): Row Data.
            cell_obj (CellObj): Cell Object.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_row(cls, sheet: XSpreadsheet, values: Row, col_start: int, row_start: int) -> None:
        """
        Inserts a row of data into spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            values (Row): Row Data.
            col_start (int): Zero-base column index.
            row_start (int): Zero-base row index.

        Returns:
            None:
        """
        ...

    @classmethod
    def set_row(cls, *args, **kwargs) -> None:
        """
        Inserts a row of data into spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            values (Row): Row Data.
            cell_obj (CellObj): Cell Object.
            cell_name (str): Name of Cell to begin the insert such as ``A1``.
            col_start (int): Zero-base column index.
            row_start (int): Zero-base row index.

        Raises:
            MissingInterfaceError: if unable to obtain interface.

        Returns:
            None:
        """
        ordered_keys = (1, 2, 3, 4)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("sheet", "values", "cell_name", "cell_obj", "col_start", "row_start")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("set_row() got an unexpected keyword argument")
            ka[1] = kwargs.get("sheet", None)
            ka[2] = kwargs.get("values", None)
            keys = ("cell_name", "cell_obj", "col_start")
            for key in keys:
                if key in kwargs:
                    ka[3] = kwargs[key]
                    break
            if count == 3:
                return ka
            ka[4] = kwargs.get("row_start", None)
            return ka

        if count not in (3, 4):
            raise TypeError("set_row() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 3:
            pos = cls.get_cell_position(str(kargs[3]))
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
        cr_data = mLo.Lo.qi(XCellRangeData, cell_range, raise_err=True)
        cr_data.setDataArray(mTblHelper.TableHelper.to_2d_tuple(values))  #  1-row 2D array

    # endregion set_row()

    # region get_row()

    @overload
    @classmethod
    def get_row(cls, cell_range: XCellRange) -> Row:
        """
        Gets a row of data from spreadsheet.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell range to get column data from.

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, ``None``.
        """
        ...

    @overload
    @classmethod
    def get_row(cls, sheet: XSpreadsheet, row_idx: int) -> Row:
        """
        Gets a row of data from spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            row_idx (int): Zero base row index such as `0` for row ``1``

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, ``None``.
        """
        ...

    @overload
    @classmethod
    def get_row(cls, sheet: XSpreadsheet, range_name: str) -> Row:
        """
        Gets a row of data from spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range such as ``A1:A12``.

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, ``None``.
        """
        ...

    @overload
    @classmethod
    def get_row(cls, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj) -> Row:
        """
        Gets a row of data from spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_obj (CellObj): Cell Object.

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, ``None``.
        """
        ...

    @overload
    @classmethod
    def get_row(cls, sheet: XSpreadsheet, range_obj: mRngObj.RangeObj) -> Row:
        """
        Gets a row of data from spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_obj (RangeObj): Range Object.

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, ``None``.
        """
        ...

    @classmethod
    def get_row(cls, *args, **kwargs) -> Row:
        """
        Gets a row of data from spreadsheet.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell range to get column data from.
            row_idx (int): Zero base row index such as `0` for row ``1``
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range such as ``A1:A12``.
            cell_obj (CellObj): Cell Object.
            range_obj (RangeObj): Range Object.

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, ``None``.
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("cell_range", "sheet", "range_name", "cell_obj", "range_obj", "row_idx")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_row() got an unexpected keyword argument")
            keys = ("cell_range", "sheet")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            keys = ("range_name", "cell_obj", "range_obj", "row_idx")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            return ka

        if count not in (1, 2):
            raise TypeError("get_row() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            vals = cls.get_array(kargs[1])
        else:  # count == 2
            sheet = cast(XSpreadsheet, kargs[1])
            row = -1
            arg2 = kargs[2]
            if isinstance(arg2, mCellObj.CellObj):
                row = arg2.row - 1
            elif isinstance(arg2, mRngObj.RangeObj):
                row = arg2.cell_start.row - 1
            elif isinstance(arg2, int):
                row = arg2
                if row < 0:
                    # there can't be a negative row.
                    return []
            if row > -1:
                used_range = cls.find_used_range(sheet)
                ca = cls._get_address_cell(used_range)
                if ca.StartRow > row or ca.EndRow < row:
                    # the requested row is outside the used area of sheet.
                    return []
                range_name = f"{cls._get_cell_str_col_row(col=ca.StartColumn, row=row)}:{cls._get_cell_str_col_row(col=ca.EndColumn, row=row)}"
                row_range = used_range.getCellRangeByName(range_name)
                vals = cls.get_array(row_range)
            else:
                vals = cast(Table, cls.get_array(sheet=sheet, range_name=str(arg2)))
        return cls.extract_row(vals=vals, row_idx=0)

    # endregion get_row()

    @staticmethod
    def extract_row(vals: Table, row_idx: int) -> Row:
        """
        Extracts a row from a table.

        |lo_safe|

        Args:
            vals (Table): Table of data.
            row_idx (int): Row index to extract.

        Raises:
            IndexError: If row_idx is out of range.

        Returns:
            Row: Row of data
        """
        row_len = len(vals)
        if row_idx < 0 or row_idx > row_len - 1:
            raise IndexError("Row index out of range")

        return vals[row_idx]

    # region get_col()
    @overload
    @classmethod
    def get_col(cls, cell_range: XCellRange) -> List[Any]:
        """
        Gets a column of data from spreadsheet.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell range to get column data from.

        Returns:
            List[Any]: 1-Dimensional List.
        """
        ...

    @overload
    @classmethod
    def get_col(cls, sheet: XSpreadsheet, col_name: str) -> List[Any]:
        """
        Gets a column of data from spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            col_name (str): column name such as ``A``.

        Returns:
            List[Any]: 1-Dimensional List.
        """
        ...

    @overload
    @classmethod
    def get_col(cls, sheet: XSpreadsheet, col_idx: int) -> List[Any]:
        """
        Gets a column of data from spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            col_idx (int): Zero base column index such as `0` for column ``A``.

        Returns:
            List[Any]: 1-Dimensional List.
        """
        ...

    @overload
    @classmethod
    def get_col(cls, sheet: XSpreadsheet, range_name: str) -> List[Any]:
        """
        Gets a column of data from spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range such as ``A1:A12``.

        Returns:
            List[Any]: 1-Dimensional List.
        """
        ...

    @overload
    @classmethod
    def get_col(cls, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj) -> List[Any]:
        """
        Gets a column of data from spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_obj (CellObj): Cell Object.

        Returns:
            List[Any]: 1-Dimensional List.
        """
        ...

    @overload
    @classmethod
    def get_col(cls, sheet: XSpreadsheet, range_obj: mRngObj.RangeObj) -> List[Any]:
        """
        Gets a column of data from spreadsheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_obj (RangeObj): Range Object.

        Returns:
            List[Any]: 1-Dimensional List.
        """
        ...

    @classmethod
    def get_col(cls, *args, **kwargs) -> List[Any]:
        """
        Gets a column of data from spreadsheet.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell range to get column data from.
            sheet (XSpreadsheet): Spreadsheet.
            col_name (str): column name such as ``A``.
            col_idx (int): Zero base column index such as `0` for column ``A``.
            range_name (str): Range such as ``A1:A12``.
            range_obj (RangeObj): Range Object.
            cell_obj (CellObj): Cell Object.

        Returns:
            List[Any]: 1-Dimensional List.
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("cell_range", "sheet", "range_name", "range_obj", "cell_obj", "col_name", "col_idx")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_col() got an unexpected keyword argument")
            keys = ("cell_range", "sheet")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            keys = ("range_name", "range_obj", "cell_obj", "col_name", "col_idx")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            return ka

        if count not in (1, 2):
            raise TypeError("get_col() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            vals = cls.get_array(kargs[1])
        else:  # count == 2
            sheet = cast(XSpreadsheet, kargs[1])
            col = -1
            arg2 = kargs[2]
            if isinstance(arg2, mCellObj.CellObj):
                col = arg2.col_obj.index
            elif isinstance(arg2, mRngObj.RangeObj):
                col = arg2.cell_start.col_obj.index
            elif isinstance(arg2, int):
                col = arg2
                if col < 0:
                    # there can't be a negative column.
                    return []
            else:
                name = str(arg2)
                if name.isalpha():
                    col = cls.column_string_to_number(name)
            if col > -1:
                used_range = cls.find_used_range(sheet)
                ca = cls._get_address_cell(used_range)
                if ca.StartColumn > col or ca.EndColumn < col:
                    # the requested col is outside the used area of sheet.
                    return []
                range_name = cls.get_range_str(col_start=col, row_start=ca.StartRow, col_end=col, row_end=ca.EndRow)
                col_range = used_range.getCellRangeByName(range_name)
                vals = cls.get_array(col_range)
            else:
                vals = cls.get_array(sheet=sheet, range_name=str(arg2))
        return cls.extract_col(vals=vals, col_idx=0)

    # endregion get_col()

    @staticmethod
    def extract_col(vals: Table, col_idx: int) -> List[Any]:
        """
        Extract column data and returns as a list.

        |lo_safe|

        Args:
            vals (Table): 2-d table of data.
            col_idx (int): column index to extract.

        Returns:
            List[Any]: Column data if found; Otherwise, empty list.
        """
        col_vals = []
        row_len = len(vals)
        if row_len == 0:
            return col_vals
        col_len = len(vals[0])
        if col_idx < 0 or col_idx > col_len - 1:
            mLo.Lo.print("Column index out of range")
            return col_vals

        col_vals.extend(row[col_idx] for row in vals)
        return col_vals

    @classmethod
    def get_col_used_first_index(cls, sheet: XSpreadsheet) -> int:
        """
        Gets the index of the column of the left edge of the used sheet range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet

        Returns:
            int: Zero based index of first column used on the sheet.
        """
        used_range = cls.find_used_range(sheet)
        ca = cls._get_address_cell(used_range)
        return ca.StartColumn

    @classmethod
    def get_col_used_last_index(cls, sheet: XSpreadsheet) -> int:
        """
        Gets the index of the column of the right edge of the used sheet range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.

        Returns:
            int: Zero based index of last column used on the sheet.
        """
        used_range = cls.find_used_range(sheet)
        ca = cls._get_address_cell(used_range)
        return ca.EndColumn

    @classmethod
    def get_row_used_first_index(cls, sheet: XSpreadsheet) -> int:
        """
        Gets the index of the row of the top edge of the used sheet range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.

        Returns:
            int: Zero based index of first row used on the sheet.
        """
        used_range = cls.find_used_range(sheet)
        ca = cls._get_address_cell(used_range)
        return ca.StartRow

    @classmethod
    def get_row_used_last_index(cls, sheet: XSpreadsheet) -> int:
        """
        Gets the index of the row of the bottom edge of the used sheet range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.

        Returns:
            int: Zero based index of last row used on the sheet.
        """
        used_range = cls.find_used_range(sheet)
        ca = cls._get_address_cell(used_range)
        return ca.EndRow

    # endregion --------------- set/get rows and columns -----------------

    # region --------------- special cell types ------------------------

    @classmethod
    def set_date(cls, sheet: XSpreadsheet, cell_name: str | mCellObj.CellObj, day: int, month: int, year: int) -> None:
        """
        Writes a date with standard date format into a spreadsheet.

        |lo_unsafe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_name (str | CellObj): Cell name.
            day (int): Date day part.
            month (int): Date month part.
            year (int): Date year part.
        """
        xcell = cls.get_cell(sheet, cell_name)
        xcell.setFormula(f"{month}/{day}/{year}")

        nfs_supplier = mLo.Lo.create_instance_mcf(XNumberFormatsSupplier, "com.sun.star.util.NumberFormatsSupplier")
        if nfs_supplier is None:
            return
        number_formats = nfs_supplier.getNumberFormats()
        format_types = mLo.Lo.qi(XNumberFormatTypes, number_formats)
        if format_types is None:
            return
        locale = Locale()
        # aLocale.Country = "GB"
        # aLocale.Language = "en"

        nformat = format_types.getStandardFormat(NumberFormat.DATE, locale)
        mProps.Props.set(xcell, NumberFormat=nformat)

    # region    add_annotation()
    @overload
    @classmethod
    def add_annotation(cls, sheet: XSpreadsheet, cell_name: str, msg: str) -> XSheetAnnotation:
        """
        Adds an annotation to a cell and makes the annotation visible.

        |lo_unsafe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_name (str): Name of cell to add annotation such as ``A1``.
            msg (str): Annotation Text.

        Raises:
            MissingInterfaceError: If interface is missing.

        Returns:
            XSheetAnnotation: Cell annotation that was added.
        """
        ...

    @overload
    @classmethod
    def add_annotation(cls, sheet: XSpreadsheet, cell_name: str, msg: str, is_visible: bool) -> XSheetAnnotation:
        """
        Adds an annotation to a cell.

        |lo_unsafe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_name (str): Name of cell to add annotation such as ``A1``.
            msg (str): Annotation Text.
            set_visible (bool): Determines if the annotation is set visible.

        Raises:
            MissingInterfaceError: If interface is missing.

        Returns:
            XSheetAnnotation: Cell annotation that was added.
        """
        ...

    @classmethod
    def add_annotation(cls, sheet: XSpreadsheet, cell_name: str, msg: str, is_visible=True) -> XSheetAnnotation:
        """
        Adds an annotation to a cell and makes the annotation visible.

        |lo_unsafe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_name (str): Name of cell to add annotation such as ``A1``.
            msg (str): Annotation Text.
            set_visible (bool): Determines if the annotation is set visible.

        Raises:
            MissingInterfaceError: If interface is missing.

        Returns:
            XSheetAnnotation: Cell annotation that was added.
        """
        # add the annotation
        addr = cls.get_cell_address(sheet=sheet, cell_name=cell_name)
        annotation_supp = mLo.Lo.qi(XSheetAnnotationsSupplier, sheet, True)
        annotation = annotation_supp.getAnnotations()
        annotation.insertNew(addr, msg)

        # get a reference to the annotation
        xcell = cls.get_cell(sheet=sheet, cell_name=cell_name)
        ann_anchor = mLo.Lo.qi(XSheetAnnotationAnchor, xcell, True)
        ann = ann_anchor.getAnnotation()
        ann.setIsVisible(is_visible)
        return ann

    # endregion add_annotation()

    @classmethod
    def get_annotation(cls, sheet: XSpreadsheet, cell_name: str | mCellObj.CellObj) -> XSheetAnnotation:
        """
        Gets an annotation of a cell.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cell_name (str | CellObj): Cell name

        Raises:
            MissingInterfaceError: If interface is missing

        Returns:
            XSheetAnnotation: Cell annotation on success; Otherwise, None
        """
        # get a reference to the annotation
        xcell = cls.get_cell(sheet, cell_name)
        ann_anchor = mLo.Lo.qi(XSheetAnnotationAnchor, xcell)
        if ann_anchor is None:
            raise mEx.MissingInterfaceError(XSheetAnnotationAnchor, f"No XSheetAnnotationAnchor for {cell_name}")
        return ann_anchor.getAnnotation()

    @classmethod
    def get_annotation_str(cls, sheet: XSpreadsheet, cell_name: str | mCellObj.CellObj) -> str:
        """
        Gets text of an annotation for a cell.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cell_name (str | CellObj): Cell name

        Returns:
            str: Cell annotation text
        """
        ann = cls.get_annotation(sheet, cell_name)
        if ann is None:
            return ""
        simple_text = mLo.Lo.qi(XSimpleText, ann)
        return "" if simple_text is None else simple_text.getString()

    # endregion ------------ special cell types ------------------------

    # region --------------- get XCell and XCellRange methods ----------

    # region    get_cell()
    @classmethod
    def _get_cell_sheet_col_row(cls, sheet: XSpreadsheet, col: int, row: int) -> XCell:
        """LO Safe Method"""
        return sheet.getCellByPosition(col, row)

    @classmethod
    def _get_cell_sheet_addr(cls, sheet: XSpreadsheet, addr: CellAddress) -> XCell:
        """LO Safe Method"""
        # not using Sheet value in addr
        return cls._get_cell_sheet_col_row(sheet=sheet, col=addr.Column, row=addr.Row)

    @classmethod
    def _get_cell_sheet_cell(cls, sheet: XSpreadsheet, cell_name: str) -> XCell:
        """LO Safe Method"""
        cell_range = sheet.getCellRangeByName(cell_name)
        return cls._get_cell_cell_rng(cell_range=cell_range, col=0, row=0)

    @classmethod
    def _get_cell_cell_rng(cls, cell_range: XCellRange, col: int, row: int) -> XCell:
        """LO Safe Method"""
        return cell_range.getCellByPosition(col, row)

    @overload
    @classmethod
    def get_cell(cls, cell: XCell) -> XCell:
        """
        Gets a cell.

        |lo_safe|

        Args:
            cell (XCell): Cell.

        Returns:
            XCell: cell.
        """
        ...

    @overload
    @classmethod
    def get_cell(cls, sheet: XSpreadsheet, addr: CellAddress) -> XCell:
        """
        Gets a cell.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            addr (CellAddress): Cell Address.

        Returns:
            XCell: cell.
        """
        ...

    @overload
    @classmethod
    def get_cell(cls, sheet: XSpreadsheet, cell_name: str) -> XCell:
        """
        Gets a cell.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_name (str): Cell Name such as ``A1``.

        Returns:
            XCell: cell.
        """
        ...

    @overload
    @classmethod
    def get_cell(cls, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj) -> XCell:
        """
        Gets a cell.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_obj: (CellObj): Cell object.

        Returns:
            XCell: cell.
        """
        ...

    @overload
    @classmethod
    def get_cell(cls, sheet: XSpreadsheet, col: int, row: int) -> XCell:
        """
        Gets a cell.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            col (int): Cell column.
            row (int): cell row.

        Returns:
            XCell: cell.
        """
        ...

    @overload
    @classmethod
    def get_cell(cls, cell_range: XCellRange) -> XCell:
        """
        Gets a cell.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range.

        Returns:
            XCell: cell.
        """
        ...

    @overload
    @classmethod
    def get_cell(cls, cell_range: XCellRange, col: int, row: int) -> XCell:
        """
        Gets a cell.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range.
            col (int): Cell column.
            row (int): cell row.

        Returns:
            XCell: cell.
        """
        ...

    @classmethod
    def get_cell(cls, *args, **kwargs) -> XCell:
        """
        Gets a cell.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            addr (CellAddress): Cell Address.
            cell_name (str): Cell Name such as ``A1``.
            cell_obj: (CellObj): Cell object.
            cell_range (XCellRange): Cell Range.
            col (int): Cell column.
            row (int): cell row.
            cell (XCell): Cell.

        Returns:
            XCell: cell.

        Note:
            if ``cell`` is passed, it is returned verbatim.

        .. versionchanged:: 0.10.0
            Added overload for ``cell`` argument.
        """
        cell = kwargs.pop("cell", None)
        if cell is not None:
            return cell

        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("sheet", "cell_range", "addr", "col", "cell_name", "cell_obj", "row")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_cell() got an unexpected keyword argument")
            keys = ("sheet", "cell_range")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            keys = ("addr", "col", "cell_name", "cell_obj")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("row", None)
            return ka

        if count not in (1, 2, 3):
            raise TypeError("get_cell() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            cell_obj = kargs[1]
            if mLo.Lo.is_uno_interfaces(cell_obj, XCell):
                return cell_obj
            # get_cell(cell_range: XCellRange)
            # cell range is relative position.
            # if a range is C4:E9 then Cell range at col=0 ,row=0 is C4
            return cls._get_cell_cell_rng(cell_range=cell_obj, col=0, row=0)

        elif count == 2:
            if isinstance(kargs[2], (str, mCellObj.CellObj)):
                # get_cell(sheet: XSpreadsheet, cell_name: str)
                return cls._get_cell_sheet_cell(sheet=kargs[1], cell_name=str(kargs[2]))
            else:
                # get_cell(sheet: XSpreadsheet, addr: CellAddress)
                return cls._get_cell_sheet_addr(sheet=kargs[1], addr=kargs[2])
        else:
            sheet = mLo.Lo.qi(XSpreadsheet, kargs[1])
            if sheet is None:
                # get_cell(cell_range: XCellRange, col: int, row: int)
                return cls._get_cell_cell_rng(cell_range=kargs[1], col=kargs[2], row=kargs[3])
            else:
                # get_cell(sheet: XSpreadsheet, col: int, row: int)
                return cls._get_cell_sheet_col_row(sheet=sheet, col=kargs[2], row=kargs[3])

    # endregion get_cell()

    @staticmethod
    def is_cell_range_name(s: str) -> bool:
        """
        Gets if is a cell name or a cell range.

        |lo_safe|

        Args:
            s (str): cell name such as 'A1' or range name such as 'B3:E7'

        Returns:
            bool: True if range name; Otherwise, False
        """
        return ":" in s

    @staticmethod
    def is_single_cell_range(cr_addr: CellRangeAddress) -> bool:
        """
        Gets if a cell address is a single cell or a range.

        |lo_safe|

        Args:
            cr_addr (CellRangeAddress): cell range address

        Returns:
            bool: ``True`` if single cell; Otherwise, ``False``
        """
        return cr_addr.StartColumn == cr_addr.EndColumn and cr_addr.StartRow == cr_addr.EndRow

    @staticmethod
    def is_single_column_range(cr_addr: CellRangeAddress) -> bool:
        """
        Gets if a cell address is a single column or multi-column.

        |lo_safe|

        Args:
            cr_addr (CellRangeAddress): cell range address

        Returns:
            bool: ``True`` if single column; Otherwise, ``False``

        Note:
            If ``cr_addr`` is a single cell address then ``True`` is returned.

        .. versionadded:: 0.8.2
        """
        return cr_addr.StartColumn == cr_addr.EndColumn

    @staticmethod
    def is_single_row_range(cr_addr: CellRangeAddress) -> bool:
        """
        Gets if a cell address is a single row or multi-row.

        |lo_safe|

        Args:
            cr_addr (CellRangeAddress): cell range address

        Returns:
            bool: ``True`` if single row; Otherwise, ``False``

        Note:
            If ``cr_addr`` is a single cell address then ``True`` is returned.

        .. versionadded:: 0.8.2
        """
        return cr_addr.StartRow == cr_addr.EndRow

    # region    get_cell_range()
    @classmethod
    def _get_cell_range_addr(cls, sheet: XSpreadsheet, addr: CellRangeAddress) -> XCellRange:
        """LO Safe Method"""
        return cls._get_cell_range_col_row(
            sheet=sheet,
            start_col=addr.StartColumn,
            start_row=addr.StartRow,
            end_col=addr.EndColumn,
            end_row=addr.EndRow,
        )

    @staticmethod
    def _get_cell_range_rng_name(sheet: XSpreadsheet, range_name: str) -> XCellRange:
        """LO Safe Method"""
        cell_range = sheet.getCellRangeByName(range_name)
        if cell_range is None:
            raise Exception(f"Could not access cell range: {range_name}")
        return cell_range

    @staticmethod
    def _get_cell_range_col_row(
        sheet: XSpreadsheet, start_col: int, start_row: int, end_col: int, end_row: int
    ) -> XCellRange:
        """LO Safe Method"""
        if start_col > end_col:
            # swap
            start_col, end_col = end_col, start_col
        if start_row > end_row:
            # swap
            start_row, end_row = end_row, start_row
        try:
            cell_range = sheet.getCellRangeByPosition(start_col, start_row, end_col, end_row)
            if cell_range is None:
                raise Exception
            return cell_range
        except Exception as e:
            raise Exception(f"Could not access cell range : ({start_col}, {start_row}, {end_col}, {end_row})") from e

    @overload
    @classmethod
    def get_cell_range(cls, sheet: XSpreadsheet, range_name: str) -> XCellRange:
        """
        Gets a cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_name (str): Range Name such as ``A1:D5``.

        Returns:
            XCellRange: Cell range
        """
        ...

    @overload
    @classmethod
    def get_cell_range(cls, sheet: XSpreadsheet, range_obj: mRngObj.RangeObj) -> XCellRange:
        """
        Gets a cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_obj (RangeObj): Range Object.

        Returns:
            XCellRange: Cell range
        """
        ...

    @overload
    @classmethod
    def get_cell_range(cls, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj) -> XCellRange:
        """
        Gets a cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            cell_obj (CellObj): Cell Object.

        Returns:
            XCellRange: Cell range
        """
        ...

    @overload
    @classmethod
    def get_cell_range(cls, sheet: XSpreadsheet, cr_addr: CellRangeAddress) -> XCellRange:
        """
        Gets a cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            cr_addr (CellRangeAddress): Cell range Address.

        Returns:
            XCellRange: Cell range
        """
        ...

    @overload
    @classmethod
    def get_cell_range(cls, cell_range: XCellRange) -> XCellRange:
        """
        Gets a cell range.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range. If passed in then the same instance is returned.

        Returns:
            XCellRange: Cell range
        """
        ...

    @overload
    @classmethod
    def get_cell_range(
        cls, sheet: XSpreadsheet, col_start: int, row_start: int, col_end: int, row_end: int
    ) -> XCellRange:
        """
        Gets a cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.

        Returns:
            XCellRange: Cell range
        """
        ...

    @classmethod
    def get_cell_range(cls, *args, **kwargs) -> XCellRange:
        """
        Gets a cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_name (str): Range Name such as ``A1:D5``.
            range_obj (RangeObj): Range Object.
            cell_obj (CellObj): Cell Object.
            cr_addr (CellRangeAddress): Cell range Address.
            cell_range (XCellRange): Cell Range. If passed in then the same instance is returned.
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.

        Raises:
            Exception: if unable to access cell range.

        Returns:
            XCellRange: Cell range
        """
        cell_rng = kwargs.pop("cell_range", None)
        if cell_rng is not None:
            try:
                return mLo.Lo.qi(cell_rng, XCellRange, True)
            except Exception as e:
                raise TypeError(f"Expected XCellRange but got {type(cell_rng).__name__}") from e

        ordered_keys = (1, 2, 3, 4, 5)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            # start_col, start_row, end_col, end_row are for backwards compatibility, Changed around ver 0.6

            valid_keys = (
                "sheet",
                "cr_addr",
                "range_name",
                "range_obj",
                "cell_obj",
                "start_col",
                "col_start",
                "start_row",
                "row_start",
                "end_col",
                "col_end",
                "end_row",
                "row_end",
            )
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_cell_range() got an unexpected keyword argument")
            keys = ("sheet",)
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            keys = ("cr_addr", "range_name", "range_obj", "cell_obj", "start_col", "col_start")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("start_row", None) if kwargs.get("row_start", None) is None else kwargs.get("row_start")
            ka[4] = kwargs.get("end_col", None) if kwargs.get("col_end", None) is None else kwargs.get("col_end")
            ka[5] = kwargs.get("end_row", None) if kwargs.get("row_end", None) is None else kwargs.get("row_end")
            return ka

        if count not in (1, 2, 5):
            raise TypeError("get_cell_range() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            # can only be: get_cell_range(cls, cell_range: XCellRange) -> XCellRange:
            try:
                return mLo.Lo.qi(kargs[1], XCellRange, True)
            except Exception as e:
                raise TypeError(f"Expected XCellRange but got {type(kargs[1]).__name__}") from e

        arg1 = cast(XSpreadsheet, kargs[1])
        arg2 = kargs[2]
        if count == 2:
            if isinstance(arg2, str):
                # def get_cell_range(sheet: XSpreadsheet, range_name: str)
                return cls._get_cell_range_rng_name(sheet=arg1, range_name=cls.get_safe_rng_str(arg2, True))
            elif isinstance(arg2, mRngObj.RangeObj):
                return cls._get_cell_range_rng_name(sheet=arg1, range_name=str(arg2))
            elif isinstance(arg2, mCellObj.CellObj):
                if arg2.range_obj:
                    return cls._get_cell_range_rng_name(sheet=arg1, range_name=str(arg2.range_obj))
                return cls._get_cell_range_rng_name(sheet=arg1, range_name=str(arg2.get_range_obj()))
            else:
                # get_cell_range(sheet: XSpreadsheet, addr:CellRangeAddress)
                return cls._get_cell_range_addr(sheet=arg1, addr=arg2)
        else:
            # get_cell_range(sheet: XSpreadsheet, start_col: int, start_row: int, end_col: int, end_row: int)
            return cls._get_cell_range_col_row(
                sheet=arg1,
                start_col=arg2,
                start_row=kargs[3],
                end_col=kargs[4],
                end_row=kargs[5],
            )

    # endregion get_cell_range()

    # region    find_used_range()

    @overload
    @classmethod
    def find_used_range(cls, sheet: XSpreadsheet) -> XCellRange:
        """
        Find used range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.

        Returns:
            XCellRange: Cell range.
        """
        ...

    @overload
    @classmethod
    def find_used_range(cls, sheet: XSpreadsheet, range_name: str) -> XCellRange:
        """
        Find used range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_name (str): Range Name such as ``A1:D5``.

        Returns:
            XCellRange: Cell range.
        """
        ...

    @overload
    @classmethod
    def find_used_range(cls, sheet: XSpreadsheet, range_obj: mRngObj.RangeObj) -> XCellRange:
        """
        Find used range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_obj (RangeObj): Range Object.

        Returns:
            XCellRange: Cell range.
        """
        ...

    @overload
    @classmethod
    def find_used_range(cls, sheet: XSpreadsheet, cr_addr: CellRangeAddress) -> XCellRange:
        """
        Find used range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            cr_addr (CellRangeAddress): Cell range Address.

        Returns:
            XCellRange: Cell range.
        """
        ...

    @classmethod
    def find_used_range(cls, *args, **kwargs) -> XCellRange:
        """
        Find used range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_name (str): Range Name such as ``A1:D5``.
            range_obj (RangeObj): Range Object.
            cr_addr (CellRangeAddress): Cell range Address.

        Returns:
            XCellRange: Cell range.

        See Also:
            - :ref:`ch20_finding_with_cursors`
        """
        # cell_name is for backwards compatibility
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("sheet", "cell_name", "range_name", "range_obj", "cr_addr")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("find_used_range() got an unexpected keyword argument")
            ka[1] = kwargs.get("sheet", None)
            if count == 1:
                return ka
            keys = ("cell_name", "range_name", "range_obj", "cr_addr")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            return ka

        if count not in (1, 2):
            raise TypeError("find_used_range() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        sheet = cast(XSpreadsheet, kargs[1])
        arg2 = kargs.get(2)

        if arg2 is None:
            cursor = sheet.createCursor()
        else:
            xrange = cls.get_cell_range(sheet, arg2)
            cell_range = mLo.Lo.qi(XSheetCellRange, xrange, True)
            cursor = sheet.createCursorByRange(cell_range)
        return cls.find_used_cursor(cursor)

    # endregion find_used_range()

    # region find_used_range_obj()

    @overload
    @classmethod
    def find_used_range_obj(cls, sheet: XSpreadsheet) -> mRngObj.RangeObj:
        """
        Find used range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    @classmethod
    def find_used_range_obj(cls, sheet: XSpreadsheet, range_name: str) -> mRngObj.RangeObj:
        """
        Find used range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_name (str): Range Name such as ``A1:D5``.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    @classmethod
    def find_used_range_obj(cls, sheet: XSpreadsheet, range_obj: mRngObj.RangeObj) -> mRngObj.RangeObj:
        """
        Find used range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_obj (RangeObj): Range Object.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    @classmethod
    def find_used_range_obj(cls, sheet: XSpreadsheet, cr_addr: CellRangeAddress) -> mRngObj.RangeObj:
        """
        Find used range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            cr_addr (CellRangeAddress): Cell range Address.

        Returns:
            RangeObj: Range object.
        """
        ...

    @classmethod
    def find_used_range_obj(cls, *args, **kwargs) -> mRngObj.RangeObj:
        """
        Find used range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_name (str): Range Name such as ``A1:D5``.
            range_obj (RangeObj): Range Object.
            cr_addr (CellRangeAddress): Cell range Address.

        Returns:
            RangeObj: Range object.

        .. versionadded:: 0.9.0
        """
        used_range = cls.find_used_range(*args, **kwargs)
        ca = cls._get_address_cell(used_range)
        return mRngObj.RangeObj.from_range(ca)

    # endregion find_used_range_obj()

    @staticmethod
    def find_used_cursor(cursor: XSheetCellCursor) -> XCellRange:
        """
        Find used cursor.

        |lo_safe|

        Args:
            cursor (XSheetCellCursor): Sheet Cursor

        Raises:
            MissingInterfaceError: if unable to find interface

        Returns:
            XCellRange: Cell range
        """
        # find the used area
        ua_cursor = mLo.Lo.qi(XUsedAreaCursor, cursor, True)
        ua_cursor.gotoStartOfUsedArea(False)
        ua_cursor.gotoEndOfUsedArea(True)

        return mLo.Lo.qi(XCellRange, ua_cursor, True)

    @staticmethod
    def get_col_range(sheet: XSpreadsheet, idx: int) -> XCellRange:
        """
        Get Column by index.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            idx (int): Zero-based column index.

        Raises:
            MissingInterfaceError: if unable to find interface.

        Returns:
            XCellRange: Cell range.
        """
        cr_range = mLo.Lo.qi(XColumnRowRange, sheet)
        if cr_range is None:
            raise mEx.MissingInterfaceError(XColumnRowRange)
        cols = cr_range.getColumns()
        con = mLo.Lo.qi(XIndexAccess, cols)
        if con is None:
            raise mEx.MissingInterfaceError(XIndexAccess)
        cell_range = mLo.Lo.qi(XCellRange, con.getByIndex(idx))
        if cell_range is None:
            raise mEx.MissingInterfaceError(XCellRange, f"Could not access range for column position: {idx}")
        return cell_range

    @staticmethod
    def get_row_range(sheet: XSpreadsheet, idx: int) -> XCellRange:
        """
        Get Row by index.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            idx (int): Zero-based column index.

        Raises:
            MissingInterfaceError: if unable to find interface.

        Returns:
            XCellRange: Cell range.
        """
        cr_range = mLo.Lo.qi(XColumnRowRange, sheet)
        if cr_range is None:
            raise mEx.MissingInterfaceError(XColumnRowRange)
        rows = cr_range.getRows()
        con = con = mLo.Lo.qi(XIndexAccess, rows)
        if con is None:
            raise mEx.MissingInterfaceError(XIndexAccess)
        cell_range = mLo.Lo.qi(XCellRange, con.getByIndex(idx))
        if cell_range is None:
            raise mEx.MissingInterfaceError(XCellRange, f"Could not access range for row position: {idx}")
        return cell_range

    # region get_cell_protection()
    @overload
    @classmethod
    def get_cell_protection(cls, cell: XCell) -> CellProtection:
        """
        Gets cell protection.

        |lo_safe|

        Args:
            cell (XCell): Cell.

        Returns:
            com.sun.star.util.CellProtection: Cell Protection.
        """
        ...

    @overload
    @classmethod
    def get_cell_protection(cls, sheet: XSpreadsheet, addr: CellAddress) -> CellProtection:
        """
        Gets cell protection.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            addr (CellAddress): Cell Address.

        Returns:
            com.sun.star.util.CellProtection: Cell Protection.
        """
        ...

    @overload
    @classmethod
    def get_cell_protection(cls, sheet: XSpreadsheet, cell_name: str) -> CellProtection:
        """
        Gets cell protection.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_name (str): Cell Name such as ``A1``.

        Returns:
            com.sun.star.util.CellProtection: Cell Protection.
        """
        ...

    @overload
    @classmethod
    def get_cell_protection(cls, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj) -> CellProtection:
        """
        Gets cell protection.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_obj: (CellObj): Cell object.

        Returns:
            com.sun.star.util.CellProtection: Cell Protection.
        """
        ...

    @overload
    @classmethod
    def get_cell_protection(cls, sheet: XSpreadsheet, col: int, row: int) -> CellProtection:
        """
        Gets cell protection.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            col (int): Cell column.
            row (int): cell row.

        Returns:
            com.sun.star.util.CellProtection: Cell Protection.
        """
        ...

    @overload
    @classmethod
    def get_cell_protection(cls, cell_range: XCellRange) -> CellProtection:
        """
        Gets cell protection.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range.

        Returns:
            com.sun.star.util.CellProtection: Cell Protection.
        """
        ...

    @overload
    @classmethod
    def get_cell_protection(cls, cell_range: XCellRange, col: int, row: int) -> CellProtection:
        """
        Gets cell protection.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range.
            col (int): Cell column.
            row (int): cell row.

        Returns:
            com.sun.star.util.CellProtection: Cell Protection.
        """
        ...

    @classmethod
    def get_cell_protection(cls, *args, **kwargs) -> CellProtection:
        """
        Gets cell protection.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            addr (CellAddress): Cell Address.
            cell_name (str): Cell Name such as ``A1``.
            cell_obj: (CellObj): Cell object.
            cell_range (XCellRange): Cell Range.
            col (int): Cell column.
            row (int): cell row.
            cell (XCell): Cell.

        Returns:
            com.sun.star.util.CellProtection: Cell Protection.

        Warning:
            Cell Protection is only valid after the current sheet has been protected.

        See Also:
            - :py:meth:`~.calc.Calc.protect_sheet`
            - :ref:`help_calc_format_direct_cell_cell_protection`

        .. versionadded:: 0.10.0
        """
        cell = cls.get_cell(*args, **kwargs)
        props = mLo.Lo.qi(XPropertySet, cell, True)
        return cast("CellProtection", props.getPropertyValue("CellProtection"))

    # endregion get_cell_protection()

    # region is_cell_protected()
    @overload
    @classmethod
    def is_cell_protected(cls, cell: XCell) -> bool:
        """
        Gets whether a cell is protected.

        |lo_safe|

        Args:
            cell (XCell): Sheet Cell.

        Returns:
            bool: ``True`` if protected; Otherwise, ``False``
        """
        ...

    @overload
    @classmethod
    def is_cell_protected(cls, sheet: XSpreadsheet, addr: CellAddress) -> bool:
        """
        Gets whether a cell is protected.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            addr: (CellAddress): Cell address.

        Returns:
            bool: ``True`` if protected; Otherwise, ``False``
        """
        ...

    @overload
    @classmethod
    def is_cell_protected(cls, sheet: XSpreadsheet, cell_name: str) -> bool:
        """
        Gets whether a cell is protected.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_name: (str): Cell name.

        Returns:
            bool: ``True`` if protected; Otherwise, ``False``
        """
        ...

    @overload
    @classmethod
    def is_cell_protected(cls, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj) -> bool:
        """
        Gets whether a cell is protected.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_obj: (CellObj): Cell object.

        Returns:
            bool: ``True`` if protected; Otherwise, ``False``
        """
        ...

    @overload
    @classmethod
    def is_cell_protected(cls, sheet: XSpreadsheet, col: int, row: int) -> bool:
        """
        Gets whether a cell is protected.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            col (int): Cell column.
            row (int): cell row.

        Returns:
            bool: ``True`` if protected; Otherwise, ``False``
        """
        ...

    @overload
    @classmethod
    def is_cell_protected(cls, cell_range: XCellRange) -> bool:
        """
        Gets whether a cell is protected.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range.

        Returns:
            bool: ``True`` if protected; Otherwise, ``False``
        """
        ...

    @overload
    @classmethod
    def is_cell_protected(cls, cell_range: XCellRange, col: int, row: int) -> bool:
        """
        Gets whether a cell is protected.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range.
            col (int): Column index.
            row (int): Row index.

        Returns:
            bool: ``True`` if protected; Otherwise, ``False``
        """
        ...

    @classmethod
    def is_cell_protected(cls, *args, **kwargs) -> bool:
        """
        Gets whether a cell is protected.

        |lo_safe|

        Args:
            cell (XCell): Sheet Cell.
            sheet (XSpreadsheet): Spreadsheet.
            addr: (CellAddress): Cell address.
            cell_name: (str): Cell name.
            cell_obj: (CellObj): Cell object.
            col (int): Cell column.
            row (int): cell row.
            cell_range (XCellRange): Cell Range.

        Returns:
            bool: ``True`` if protected; Otherwise, ``False``

        Warning:
            Cell Protection is only valid after the current sheet has been protected.

        See Also:
            - :py:meth:`~.calc.Calc.protect_sheet`
            - :ref:`help_calc_format_direct_cell_cell_protection`

        .. versionadded:: 0.10.0
        """
        cp = cls.get_cell_protection(*args, **kwargs)
        return cp.IsLocked

    # endregion is_cell_protected()

    # endregion ------------ get XCell and XCellRange methods ----------

    # region --------------- convert cell/cell range names to positions -

    # region get_cell_range_positions()
    @overload
    @classmethod
    def get_cell_range_positions(cls, range_obj: mRngObj.RangeObj) -> Tuple[Point, Point]:
        """
        Gets Cell range as a tuple of Point, Point.

        - First Point.X is start column index, Point.Y is start row index.
        - Second Point.X is end column index, Point.Y is end row index.

        |lo_safe|

        Args:
            range_obj (RangeObj): Range object.

        Returns:
            Tuple[Point, Point]: Range as tuple. Point values are zero-based indexes.
        """
        ...

    @overload
    @classmethod
    def get_cell_range_positions(cls, range_values: mRngValues.RangeValues) -> Tuple[Point, Point]:
        """
        Gets Cell range as a tuple of Point, Point.

        - First Point.X is start column index, Point.Y is start row index.
        - Second Point.X is end column index, Point.Y is end row index.

        |lo_safe|

        Args:
            range_values (RangeValues): Range values.

        Returns:
            Tuple[Point, Point]: Range as tuple. Point values are zero-based indexes.
        """
        ...

    @overload
    @classmethod
    def get_cell_range_positions(cls, range_name: str) -> Tuple[Point, Point]:
        """
        Gets Cell range as a tuple of Point, Point.

        - First Point.X is start column index, Point.Y is start row index.
        - Second Point.X is end column index, Point.Y is end row index.

        |lo_safe|

        Args:
            range_name (str): Range name such as ``A1:C8``.

        Returns:
            Tuple[Point, Point]: Range as tuple. Point values are zero-based indexes.
        """
        ...

    @classmethod
    def get_cell_range_positions(cls, *args, **kwargs) -> Tuple[Point, Point]:
        """
        Gets Cell range as a tuple of Point, Point.

        - First Point.X is start column index, Point.Y is start row index.
        - Second Point.X is end column index, Point.Y is end row index.

        |lo_safe|

        Args:
            range_name (str): Range name such as ``A1:C8``.
            range_obj (RangeObj): Range object
            range_values (RangeValues): Range values

        Raises:
            ValueError: if invalid range name

        Returns:
            Tuple[Point, Point]: Range as tuple. Point values are zero-based indexes.
        """
        kargs_len = len(kwargs)
        count = len(args) + kargs_len
        if count != 1:
            raise TypeError("get_cell_range_positions() got an invalid number of arguments")

        rng = None
        for v in kwargs.values():
            rng = v

        if rng is None:
            rng = args[0]

        if isinstance(rng, str):
            rv = mRngValues.RangeValues.from_range(rng)
        elif isinstance(rng, mRngObj.RangeObj):
            rv = rng.get_range_values()
        else:
            rv = cast(mRngValues.RangeValues, rng)
        point_start = Point(rv.col_start, rv.row_start)
        point_end = Point(rv.col_end, rv.row_end)
        return (point_start, point_end)

    # endregion get_cell_range_positions()

    # region    get_cell_position()
    @classmethod
    def get_cell_position(cls, cell_name: str | mCellObj.CellObj) -> Point:
        """
        Gets a cell name as a Point.

        - ``Point.X`` is column zero-based index.
        - ``Point.Y`` is row zero-based index.

        |lo_safe|

        Args:
            cell_name (str | CellObj): Cell name

        Returns:
            Point: cell name as Point with X as col and Y as row
        """
        if isinstance(cell_name, str):
            co = mCellObj.CellObj.from_cell(cell_name)
        else:
            co = cell_name
        return Point(co.col_obj.index, co.row_obj.index)

    @classmethod
    def get_cell_pos(cls, sheet: XSpreadsheet, cell_name: str | mCellObj.CellObj) -> Point:
        """
        Contains the position of the top left cell of this range in the sheet (in 1/100 mm).

        This property contains the absolute position in the whole sheet,
        not the position in the visible area.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet
            cell_name (str | CellObj):  Cell name

        Returns:
            Point: cell name as Point
        """
        xcell = cls.get_cell(sheet, cell_name)
        pos = None
        with contextlib.suppress(mEx.PropertyNotFoundError):
            pos = mProps.Props.get(xcell, "Position")
        if pos is None:
            mLo.Lo.print(f"Could not determine position of cell '{cell_name}'")
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

        |lo_safe|

        Args:
            col_str (str):Case insensitive column name such as 'a' or 'AB'

        Returns:
            int: Zero based int representing column name
        """
        i = mTblHelper.TableHelper.col_name_to_int(name=col_str)
        return i - 1  # convert to zero based.

    @staticmethod
    def row_string_to_number(row_str: str) -> int:
        """
        Converts a string containing an int into an int.

        |lo_safe|

        Args:
            row_str (str): string to convert

        Returns:
            int: Number if conversion succeeds; Otherwise, 0
        """
        try:
            return mTblHelper.TableHelper.row_name_to_int(row_str) - 1
        except ValueError:
            mLo.Lo.print(f"Incorrect format for {row_str}")
        return 0

    # endregion ----------- convert cell/cell range names to positions --

    # region --------------- get cell and cell range addresses ---------

    # region get_safe_rng_str()
    @overload
    @classmethod
    def get_safe_rng_str(cls, range_name: str) -> str:
        """
        Gets safe range string.

        If range name is out of order then correct order is returned.

        |lo_safe|

        Args:
            range_name (str): range name such as ``A1.B7`` or ``Sheet1.A1.B7``.

        Returns:
            str: Range name as string with correct column an row order.
        """
        ...

    @overload
    @classmethod
    def get_safe_rng_str(cls, range_name: str, allow_cell_name: bool) -> str:
        """
        Gets safe range string.

        If range name is out of order then correct order is returned.

        |lo_safe|

        Args:
            range_name (str): range name such as ``A1.B7`` or ``Sheet1.A1.B7``.
            allow_cell_name: Determines if ``range_name`` accepts cell name input.

        Returns:
            str: Range name as string with correct column an row order.
        """
        ...

    @classmethod
    def get_safe_rng_str(cls, range_name: str, allow_cell_name: bool = False) -> str:
        """
        Gets safe range string.

        If range name is out of order then correct order is returned.


        For instance:

            - ``A7:B2`` returns ``A2:B7``
            - ``R7:B22`` returns ``B7:R22``

        |lo_safe|

        Args:
            range_name (str): range name such as ``A1.B7`` or ``Sheet1.A1.B7``.
            allow_cell_name: Determines if ``range_name`` accepts cell name input.

        Returns:
            str: Range name as string with correct column an row order.

        Note:
            If ``allow_cell_name`` is ``True`` and ``range_name`` is a cell name then
            the cell name is converted into a range string.

            - ``C2`` is returned as ``C2:C2``
            - ``Sheet1.C2`` is returned as ``Sheet1.C2:C2``

        .. versionadded:: 0.9.0
        """
        try:
            parts = mTblHelper.TableHelper.get_range_parts(range_name)
            return str(parts)
        except Exception:
            if not allow_cell_name:
                raise
            if cls.is_cell_range_name(range_name):
                raise
        cell = mTblHelper.TableHelper.get_cell_parts(range_name)
        # convert to a range string
        return f"{cell}:{cell.col}{cell.row}"

    # endregion get_safe_rng_str()

    # region    get_cell_address()

    @staticmethod
    def _get_cell_address_cell(cell: XCell) -> CellAddress:
        """LO Safe Method"""
        addr = mLo.Lo.qi(XCellAddressable, cell)
        if addr is None:
            raise mEx.MissingInterfaceError(XCellAddressable)
        return addr.getCellAddress()

    @classmethod
    def _get_cell_address_sheet(cls, sheet: XSpreadsheet, cell_name: str) -> CellAddress:
        """LO Safe Method"""
        cell_range = sheet.getCellRangeByName(cell_name)
        start_cell = cls._get_cell_cell_rng(cell_range=cell_range, col=0, row=0)
        return cls._get_cell_address_cell(start_cell)

    @overload
    @classmethod
    def get_cell_address(cls, cell: XCell) -> CellAddress:
        """
        Gets Cell Address.

        |lo_safe|

        Args:
            cell (XCell): Cell.

        Returns:
            CellAddress: Cell Address.
        """
        ...

    @overload
    @classmethod
    def get_cell_address(cls, sheet: XSpreadsheet, cell_name: str) -> CellAddress:
        """
        Gets Cell Address.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_name (str): Cell name such as ``A1``.

        Returns:
            CellAddress: Cell Address.
        """
        ...

    @overload
    @classmethod
    def get_cell_address(cls, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj) -> CellAddress:
        """
        Gets Cell Address.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_obj (CellObj): Cell object.

        Returns:
            CellAddress: Cell Address.
        """
        ...

    @overload
    @classmethod
    def get_cell_address(cls, sheet: XSpreadsheet, addr: CellAddress) -> CellAddress:
        """
        Gets Cell Address.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            addr (CellAddress): Cell Address.

        Returns:
            CellAddress: Cell Address.
        """
        ...

    @overload
    @classmethod
    def get_cell_address(cls, sheet: XSpreadsheet, col: int, row: int) -> CellAddress:
        """
        Gets Cell Address.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            col (int): Zero-base column index.
            row (int): Zero-base row index.

        Returns:
            CellAddress: Cell Address.
        """
        ...

    @classmethod
    def get_cell_address(cls, *args, **kwargs) -> CellAddress:
        """
        Gets Cell Address.

        |lo_safe|

        Args:
            cell (XCell): Cell.
            sheet (XSpreadsheet): Spreadsheet.
            cell_name (str): Cell name such as ``A1``.
            cell_obj (CellObj): Cell object.
            addr (CellAddress): Cell Address.
            col (int): Zero-base column index.
            row (int): Zero-base row index.

        Raises:
            MissingInterfaceError: if unable to obtain interface.

        Returns:
            CellAddress: Cell Address.
        """
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("cell", "sheet", "cell_name", "cell_obj", "col", "addr", "row")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_cell_address() got an unexpected keyword argument")
            keys = ("cell", "sheet")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            keys = ("cell_name", "cell_obj", "col", "addr")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("row", None)
            return ka

        if count not in (1, 2, 3):
            raise TypeError("get_cell_address() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            return cls._get_cell_address_cell(cell=kargs[1])
        elif count == 2:
            if isinstance(kargs[2], (str, mCellObj.CellObj)):
                return cls._get_cell_address_sheet(sheet=kargs[1], cell_name=str(kargs[2]))
            cell_name = cls._get_cell_str_addr(addr=kargs[2])
            return cls._get_cell_address_sheet(sheet=kargs[1], cell_name=cell_name)
        # elif count == 3:
        cell_name = cls._get_cell_str_col_row(col=kargs[2], row=kargs[3])
        return cls._get_cell_address_sheet(sheet=kargs[1], cell_name=cell_name)

    # endregion get_cell_address()

    # region    get_address()
    @classmethod
    def _get_address_cell(cls, cell_range: XCellRange) -> CellRangeAddress:
        """LO Safe Method"""
        addr = mLo.Lo.qi(XCellRangeAddressable, cell_range, True)
        return addr.getRangeAddress()  # type: ignore

    @classmethod
    def _get_address_sht_rng(cls, sheet: XSpreadsheet, range_name: str) -> CellRangeAddress:
        """LO Safe Method"""
        return cls._get_address_cell(cls._get_cell_range_rng_name(sheet=sheet, range_name=range_name))

    @overload
    @classmethod
    def get_address(cls, cell_range: XCellRange) -> CellRangeAddress:
        """
        Gets Range Address.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range.

        Returns:
            CellRangeAddress: Cell Range Address.
        """
        ...

    @overload
    @classmethod
    def get_address(cls, sheet: XSpreadsheet, range_name: str) -> CellRangeAddress:
        """
        Gets Range Address.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range name such as ``A1:D7``.

        Returns:
            CellRangeAddress: Cell Range Address.
        """
        ...

    @overload
    @classmethod
    def get_address(cls, sheet: XSpreadsheet, range_obj: mRngObj.RangeObj) -> CellRangeAddress:
        """
        Gets Range Address.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_obj (RangeObj): Range Object.

        Returns:
            CellRangeAddress: Cell Range Address.
        """
        ...

    @overload
    @classmethod
    def get_address(
        cls, sheet: XSpreadsheet, start_col: int, start_row: int, end_col: int, end_row: int
    ) -> CellRangeAddress:
        """
        Gets Range Address.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            start_col (int): Zero-base start column index.
            start_row (int): Zero-base start row index.
            end_col (int): Zero-base end column index.
            end_row (int): Zero-base end row index.

        Returns:
            CellRangeAddress: Cell Range Address.
        """
        ...

    @classmethod
    def get_address(cls, *args, **kwargs) -> CellRangeAddress:
        """
        Gets Range Address.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range.
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range name such as ``A1:D7``.
            range_obj (RangeObj): Range Object.
            start_col (int): Zero-base start column index.
            start_row (int): Zero-base start row index.
            end_col (int): Zero-base end column index.
            end_row (int): Zero-base end row index.

        Raises:
            MissingInterfaceError: if unable to obtain interface.

        Returns:
            CellRangeAddress: Cell Range Address.
        """
        ordered_keys = (1, 2, 3, 4, 5)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = (
                "cell_range",
                "sheet",
                "range_name",
                "range_obj",
                "start_col",
                "start_row",
                "end_col",
                "end_row",
            )
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_address() got an unexpected keyword argument")
            keys = ("cell_range", "sheet")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            keys = ("range_name", "range_obj", "start_col")
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

        if count not in (1, 2, 5):
            raise TypeError("get_address() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 1:
            # range_name or range_obj
            return cls._get_address_cell(cell_range=kargs[1])
        elif count == 2:
            arg2 = kargs[2]
            if isinstance(arg2, str):
                range_name = cls.get_safe_rng_str(arg2, True)
            else:
                range_name = str(arg2)
            return cls._get_address_sht_rng(sheet=kargs[1], range_name=range_name)
        else:
            range_name = cls._get_range_str_col_row(
                col_start=kargs[2], row_start=kargs[3], col_end=kargs[4], row_end=kargs[5]
            )
            return cls._get_address_sht_rng(sheet=kargs[1], range_name=range_name)

    # endregion get_address()

    # region    print_cell_address()
    @overload
    @classmethod
    def print_cell_address(cls, cell: XCell) -> None:
        """
        Prints Cell to console such as ``Cell: Sheet1.D3``.

        |lo_safe|

        Args:
            cell (XCell): cell.
        """
        ...

    @overload
    @classmethod
    def print_cell_address(cls, addr: CellAddress) -> None:
        """
        Prints Cell to console such as ``Cell: Sheet1.D3``.

        |lo_safe|

        Args:
            addr (CellAddress): Cell Address.
        """
        ...

    @classmethod
    def print_cell_address(cls, *args, **kwargs) -> None:
        """
        Prints Cell to console such as ``Cell: Sheet1.D3``

        |lo_safe|

        Args:
            cell (XCell): cell
            addr (CellAddress): Cell Address

        Returns:
            None:

        .. versionchanged:: 0.6.10

            Removed cancel event args.
        """
        ordered_keys = (1,)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("cell", "addr")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("print_cell_address() got an unexpected keyword argument")
            keys = ("cell", "addr")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            return ka

        if count != 1:
            raise TypeError("print_cell_address() got an invalid number of arguments")

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
    @classmethod
    def print_address(cls, cell_range: XCellRange) -> None:
        """
        Prints Cell range to console such as ``'Range: Sheet1.C3:F22``.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell range.
        """
        ...

    @overload
    @classmethod
    def print_address(cls, cr_addr: CellRangeAddress) -> None:
        """
        Prints Cell range to console such as ``'Range: Sheet1.C3:F22``.

        |lo_safe|

        Args:
            cr_addr (CellRangeAddress): Cell Address.
        """
        ...

    @classmethod
    def print_address(cls, *args, **kwargs) -> None:
        """
        Prints Cell range to console such as ``Range: Sheet1.C3:F22``.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell range.
            cr_addr (CellRangeAddress): Cell Address.

        Returns:
            None:

        .. versionchanged:: 0.6.10

            Removed cancel event args.
        """
        ordered_keys = (1,)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("cell_range", "cr_addr")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("print_address() got an unexpected keyword argument")
            keys = ("cell_range", "cr_addr")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            return ka

        if count != 1:
            raise TypeError("print_address() got an invalid number of arguments")

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
        Prints Address for one or more CellRangeAddress.

        |lo_safe|

        Returns:
            None:

        .. versionchanged:: 0.6.10

            Removed cancel event args.
        """
        print(f"No of cellrange addresses: {len(cr_addrs)}")
        for cr_addr in cr_addrs:
            cls.print_address(cr_addr=cr_addr)
        print()

    # region get_cell_series()
    @overload
    @staticmethod
    def get_cell_series(sheet: XSpreadsheet, range_name: str) -> XCellSeries:
        """
        Get cell series for a range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range name such as ``A1:B7``.

        Returns:
            XCellSeries: Cell series.
        """
        ...

    @overload
    @staticmethod
    def get_cell_series(sheet: XSpreadsheet, range_obj: mRngObj.RangeObj) -> XCellSeries:
        """
        Get cell series for a range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_obj (RangeObj): Range object.

        Returns:
            XCellSeries: Cell series.
        """
        ...

    @staticmethod
    def get_cell_series(*args, **kwargs) -> XCellSeries:
        """
        Get cell series for a range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range name such as ``A1:B7``.
            range_obj (RangeObj): Range object.

        Raises:
            MissingInterfaceError: if unable to obtain interface.

        Returns:
            XCellSeries: Cell series.

        See Also:
            :ref:`ch24_generating_data`
        """
        kargs_len = len(kwargs)
        count = len(args) + kargs_len
        ordered_keys = (1, 2)

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("sheet", "range_name", "range_obj")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_cell_series() got an unexpected keyword argument")
            ka[1] = kwargs.get("sheet", None)
            keys = ("range_name", "range_obj")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            return ka

        if count != 2:
            raise TypeError("get_cell_series() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        sheet = cast(XSpreadsheet, kargs[1])

        cell_range = sheet.getCellRangeByName(str(kargs[2]))
        series = mLo.Lo.qi(XCellSeries, cell_range, True)
        return series

    # endregion get_cell_series()

    # region    is_equal_addresses()
    @overload
    @staticmethod
    def is_equal_addresses(addr1: CellAddress, addr2: CellAddress) -> bool:
        """
        Gets if two instances of CellAddress are equal.

        |lo_safe|

        Args:
            addr1 (CellAddress): Cell Address.
            addr2 (CellAddress): Cell Address.

        Returns:
            bool: ``True`` if equal; Otherwise, ``False``.
        """
        ...

    @overload
    @staticmethod
    def is_equal_addresses(addr1: CellRangeAddress, addr2: CellRangeAddress) -> bool:
        """
        Gets if two instances of CellRangeAddress are equal.

        |lo_safe|

        Args:
            addr1 (CellRangeAddress): Cell Range Address.
            addr2 (CellRangeAddress): Cell Range Address.

        Returns:
            bool: ``True`` if equal; Otherwise, ``False``.
        """
        ...

    @staticmethod
    def is_equal_addresses(addr1: Any, addr2: Any) -> bool:
        """
        Gets if two instances of CellRangeAddress are equal.

        |lo_safe|

        Args:
            addr1 (CellAddress | CellRangeAddress): Cell address or cell range address.
            addr2 (CellAddress | CellRangeAddress): Cell address or cell range address.

        Returns:
            bool: ``True`` if equal; Otherwise, ``False``.
        """
        if addr1 is None or addr2 is None:
            return False
        try:
            is_same_type = addr1.typeName == addr2.typeName  # type: ignore
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
        """Lo Safe Method. Return as str using the name taken from the sheet works, Sheet1.A1:B2"""
        return cls._get_range_str_cr_addr_sht(cls._get_address_cell(cell_range=cell_range), sheet)

    @classmethod
    def _get_range_str_cr_addr_sht(cls, cr_addr: CellRangeAddress, sheet: XSpreadsheet) -> str:
        """LO Safe Method. Return as str using the name taken from the sheet works, Sheet1.A1:B2"""
        return f"{cls.get_sheet_name(sheet=sheet)}.{cls._get_range_str_cr_addr(cr_addr)}"

    @classmethod
    def _get_range_str_cell_rng(cls, cell_range: XCellRange) -> str:
        """Lo Safe Method. Return as str, A1:B2"""
        return cls._get_range_str_cr_addr(cls._get_address_cell(cell_range=cell_range))

    @classmethod
    def _get_range_str_cr_addr(cls, cr_addr: CellRangeAddress) -> str:
        """LO Safe Method. Return as str, A1:B2"""
        result = f"{cls._get_cell_str_col_row(cr_addr.StartColumn, cr_addr.StartRow)}:"
        result += f"{cls._get_cell_str_col_row(cr_addr.EndColumn, cr_addr.EndRow)}"
        return result

    @classmethod
    def _get_range_str_col_row(cls, col_start: int, row_start: int, col_end: int, row_end: int) -> str:
        """
        LO Safe Method.

        Return as str, A1:B2
        """
        c_start = col_start
        c_end = col_end
        r_start = row_start
        rend = row_end
        if c_start > c_end:
            # swap
            c_start, c_end = c_end, c_start
        if r_start > rend:
            # swap
            r_start, rend = rend, r_start
        return f"{cls._get_cell_str_col_row(c_start, r_start)}:{cls._get_cell_str_col_row(c_end, rend)}"

    @overload
    @classmethod
    def get_range_str(cls, cell_range: XCellRange) -> str:
        """
        Gets the range as a string in format of ``A1:B2`` or ``Sheet1.A1:B2``.

        If ``sheet`` is included the format ``Sheet1.A1:B2`` is returned; Otherwise,
        ``A1:B2`` format is returned.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range

        Returns:
            str: range as string
        """
        ...

    @overload
    @classmethod
    def get_range_str(cls, range_obj: mRngObj.RangeObj) -> str:
        """
        Gets the range as a string in format of ``A1:B2`` or ``Sheet1.A1:B2``.

        If ``sheet`` is included the format ``Sheet1.A1:B2`` is returned; Otherwise,
        ``A1:B2`` format is returned.

        |lo_safe|

        Args:
            range_obj (RangeObj): Range Object

        Returns:
            str: range as string
        """
        ...

    @overload
    @classmethod
    def get_range_str(cls, cr_addr: CellRangeAddress) -> str:
        """
        Gets the range as a string in format of ``A1:B2`` or ``Sheet1.A1:B2``.

        A ``A1:B2`` format is returned.

        |lo_safe|

        Args:
            cr_addr (CellRangeAddress): Cell Range Address

        Returns:
            str: range as string
        """
        ...

    @overload
    @classmethod
    def get_range_str(cls, cell_obj: mCellObj.CellObj) -> str:
        """
        Gets the range as a string in format of ``A1:B2`` or ``Sheet1.A1:B2``.

        A ``A1:B2`` format is returned.

        |lo_safe|

        Args:
            cell_obj (CellObj): Cell Object

        Returns:
            str: range as string
        """
        ...

    @overload
    @classmethod
    def get_range_str(cls, cell_range: XCellRange, sheet: XSpreadsheet) -> str:
        """
        Gets the range as a string in format of ``A1:B2`` or ``Sheet1.A1:B2``.

        A ``Sheet1.A1:B2`` format is returned.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range
            sheet (XSpreadsheet): Spreadsheet

        Returns:
            str: range as string
        """
        ...

    @overload
    @classmethod
    def get_range_str(cls, cr_addr: CellRangeAddress, sheet: XSpreadsheet) -> str:
        """
        Gets the range as a string in format of ``A1:B2`` or ``Sheet1.A1:B2``.

        A ``Sheet1.A1:B2`` format is returned.

        |lo_safe|

        Args:
            cr_addr (CellRangeAddress): Cell Range Address
            sheet (XSpreadsheet): Spreadsheet

        Returns:
            str: range as string
        """
        ...

    @overload
    @classmethod
    def get_range_str(cls, col_start: int, row_start: int, col_end: int, row_end: int) -> str:
        """
        Gets the range as a string in format of ``A1:B2`` or ``Sheet1.A1:B2``.

        A ``A1:B2`` format is returned.

        |lo_safe|

        Args:
            col_start (int): Zero-based start column index
            row_start (int): Zero-based start row index
            col_end (int): Zero-based end column index
            row_end (int): Zero-based end row index

        Returns:
            str: range as string
        """
        ...

    @overload
    @classmethod
    def get_range_str(cls, col_start: int, row_start: int, col_end: int, row_end: int, sheet: XSpreadsheet) -> str:
        """
        Gets the range as a string in format of ``A1:B2`` or ``Sheet1.A1:B2``.

        A ``Sheet1.A1:B2`` format is returned.

        |lo_safe|

        Args:
            col_start (int): Zero-based start column index
            row_start (int): Zero-based start row index
            col_end (int): Zero-based end column index
            row_end (int): Zero-based end row index
            sheet (XSpreadsheet): Spreadsheet

        Returns:
            str: range as string
        """
        ...

    @classmethod
    def get_range_str(cls, *args, **kwargs) -> str:
        """
        Gets the range as a string in format of ``A1:B2`` or ``Sheet1.A1:B2``.

        If ``sheet`` is included the format ``Sheet1.A1:B2`` is returned; Otherwise,
        ``A1:B2`` format is returned.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range
            range_obj (RangeObj): Range Object
            cr_addr (CellRangeAddress): Cell Range Address
            cell_obj (CellObj): Cell Object
            sheet (XSpreadsheet): Spreadsheet
            col_start (int): Zero-based start column index
            row_start (int): Zero-based start row index
            col_end (int): Zero-based end column index
            row_end (int): Zero-based end row index

        Returns:
            str: range as string
        """
        ordered_keys = (1, 2, 3, 4, 5)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            # start_col, start_row, end_col, end_row are for backward compatibility, Changed around ver 0.6
            valid_keys = (
                "cell_range",
                "range_obj",
                "cell_obj",
                "cr_addr",
                "sheet",
                "start_col",
                "col_start",
                "start_row",
                "row_start",
                "end_col",
                "col_end",
                "end_row",
                "row_end",
            )
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_range_str() got an unexpected keyword argument")
            keys = ("cell_range", "range_obj", "cell_obj", "cr_addr", "start_col", "col_start")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            if count < 5:
                keys = ("sheet", "start_row", "row_start")
            else:
                keys = ("start_row", "row_start")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("end_col", None) if kwargs.get("col_end", None) is None else kwargs.get("col_end")
            ka[4] = kwargs.get("end_row", None) if kwargs.get("row_end", None) is None else kwargs.get("row_end")
            if count == 4:
                return ka
            ka[5] = kwargs.get("sheet", None)
            return ka

        if count not in (1, 2, 4, 5):
            raise TypeError("get_range_str() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        arg1 = kargs[1]

        if count == 1:
            if isinstance(arg1, mRngObj.RangeObj):
                return arg1.to_string(True)

            if isinstance(arg1, mCellObj.CellObj):
                range_obj = arg1.get_range_obj()
                return range_obj.to_string(True)

            if mInfo.Info.is_type_interface(arg1, "com.sun.star.table.XCellRange"):
                # get_range_str(cell_range: XCellRange)
                return cls._get_range_str_cell_rng(cell_range=arg1)

            # get_range_str(cr_addr: CellRangeAddress)
            return cls._get_range_str_cr_addr(cr_addr=arg1)

        elif count == 2:
            if mInfo.Info.is_type_interface(arg1, "com.sun.star.table.XCellRange"):
                # def get_range_str(cell_range: XCellRange, sheet: XSpreadsheet)
                return cls._get_range_str_cell_rng_sht(cell_range=arg1, sheet=kargs[2])
            else:
                # get_range_str(cr_addr: CellRangeAddress, sheet: XSpreadsheet)
                return cls._get_range_str_cr_addr_sht(cr_addr=arg1, sheet=kargs[2])
        elif count == 4:
            # get_range_str(start_col:int, start_row:int, end_col:int, end_row:int)
            return cls._get_range_str_col_row(col_start=arg1, row_start=kargs[2], col_end=kargs[3], row_end=kargs[4])
        elif count == 5:
            # get_range_str(start_col: int, start_row: int, end_col: int, end_row: int,  sheet: XSpreadsheet)
            rng_str = cls._get_range_str_col_row(
                col_start=arg1, row_start=kargs[2], col_end=kargs[3], row_end=kargs[4]
            )
            return f"{cls.get_sheet_name(sheet=kargs[5])}.{rng_str}"
        return ""

    # endregion get_range_str()

    # region get_range_obj()
    @overload
    @classmethod
    def get_range_obj(cls, range_name: str) -> mRngObj.RangeObj:
        """
        Gets a range Object representing a range.

        |lo_safe|

        Args:
            range_name (str): Cell range as string.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    @classmethod
    def get_range_obj(cls, cell_range: XCellRange) -> mRngObj.RangeObj:
        """
        Gets a range Object representing a range.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    @classmethod
    def get_range_obj(cls, cr_addr: CellRangeAddress) -> mRngObj.RangeObj:
        """
        Gets a range Object representing a range.

        |lo_safe|

        Args:
            cr_addr (CellRangeAddress): Cell Range Address.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    @classmethod
    def get_range_obj(cls, range_obj: mRngObj.RangeObj) -> mRngObj.RangeObj:
        """
        Gets a range Object representing a range.

        |lo_safe|

        Args:
            range_obj (RangeObj): Range Object. If passed in the same RangeObj is returned.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    @classmethod
    def get_range_obj(cls, cell_obj: mCellObj.CellObj) -> mRngObj.RangeObj:
        """
        Gets a range Object representing a range.

        |lo_safe|

        Args:
            cell_obj (CellObj): Cell Object.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    @classmethod
    def get_range_obj(cls, cell_range: XCellRange, sheet: XSpreadsheet) -> mRngObj.RangeObj:
        """
        Gets a range Object representing a range.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range.
            sheet (XSpreadsheet): Spreadsheet.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    @classmethod
    def get_range_obj(cls, col_start: int, row_start: int, col_end: int, row_end: int) -> mRngObj.RangeObj:
        """
        Gets a range Object representing a range.

        |lo_safe|

        Args:
            col_start (int): Zero-based start column index.
            row_start (int): Zero-based start row index.
            col_end (int): Zero-based end column index.
            row_end (int): Zero-based end row index.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    @classmethod
    def get_range_obj(
        cls, col_start: int, row_start: int, col_end: int, row_end: int, sheet: XSpreadsheet
    ) -> mRngObj.RangeObj:
        """
        Gets a range Object representing a range.

        |lo_safe|

        Args:
            col_start (int): Zero-based start column index.
            row_start (int): Zero-based start row index.
            col_end (int): Zero-based end column index.
            row_end (int): Zero-based end row index.
            sheet (XSpreadsheet): Spreadsheet.

        Returns:
            RangeObj: Range object.
        """
        ...

    @classmethod
    def get_range_obj(cls, *args, **kwargs) -> mRngObj.RangeObj:
        """
        Gets a range Object representing a range.

        |lo_safe|

        Args:
            range_name (str): Cell range as string.
            cell_range (XCellRange): Cell Range.
            sheet (XSpreadsheet): Spreadsheet.
            cr_addr (CellRangeAddress): Cell Range Address.
            cell_obj (CellObj): Cell Object.
            range_obj (RangeObj): Range Object. If passed in the same RangeObj is returned.
            col_start (int): Zero-based start column index.
            row_start (int): Zero-based start row index.
            col_end (int): Zero-based end column index.
            row_end (int): Zero-based end row index.

        Returns:
            RangeObj: Range object.

        .. versionadded:: 0.8.2
        """

        def get_range(range_obj: mRngObj.RangeObj) -> mRngObj.RangeObj:

            if range_obj.sheet_idx < 0:
                return range_obj.set_sheet_index()
            return range_obj

        kargs_len = len(kwargs)
        count = len(args) + kargs_len
        if count == 1:
            val = None
            for v in kwargs.values():
                val = v
            if val is None and args:
                val = args[0]

            if val:
                if mInfo.Info.is_instance(val, mRngObj.RangeObj):
                    return get_range(val)
                if mInfo.Info.is_instance(val, str):
                    return get_range(mRngObj.RangeObj.from_range(range_val=val))
                if mInfo.Info.is_instance(val, mCellObj.CellObj):
                    return get_range(val.get_range_obj())

        range_name = cls.get_range_str(*args, **kwargs)
        return get_range(mRngObj.RangeObj.from_range(range_name))

    # endregion get_range_obj()

    # region get_range_size()
    @overload
    @classmethod
    def get_range_size(cls, range_obj: mRngObj.RangeObj) -> Size:
        """
        Gets range size.

        |lo_safe|

        Args:
            range_obj (RangeObj): Range Object.

        Returns:
            Size: Size, Width is number of Columns and Height is number of Rows.

        .. versionadded:: 0.8.2
        """
        ...

    @overload
    @classmethod
    def get_range_size(cls, cell_range: XCellRange) -> Size:
        """
        Gets range size.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range.

        Returns:
            Size: Size, Width is number of Columns and Height is number of Rows.

        .. versionadded:: 0.8.2
        """
        ...

    @overload
    @classmethod
    def get_range_size(cls, cr_addr: CellRangeAddress) -> Size:
        """
        Gets range size.

        |lo_safe|

        Args:
            cr_addr (CellRangeAddress): Cell Range Address.

        Returns:
            Size: Size, Width is number of Columns and Height is number of Rows.

        .. versionadded:: 0.8.2
        """
        ...

    @overload
    @classmethod
    def get_range_size(cls, col_start: int, row_start: int, col_end: int, row_end: int) -> Size:
        """
        Gets range size.

        |lo_safe|

        Args:
            col_start (int): Zero-based start column index.
            row_start (int): Zero-based start row index.
            col_end (int): Zero-based end column index.
            row_end (int): Zero-based end row index.

        Returns:
            Size: Size, Width is number of Columns and Height is number of Rows.

        .. versionadded:: 0.8.2
        """
        ...

    @classmethod
    def get_range_size(cls, *args, **kwargs) -> Size:
        """
        Gets range size.

        |lo_safe|

        Args:
            range_obj (RangeObj): Range Object.
            cell_range (XCellRange): Cell Range.
            cr_addr (CellRangeAddress): Cell Range Address.
            col_start (int): Zero-based start column index.
            row_start (int): Zero-based start row index.
            col_end (int): Zero-based end column index.
            row_end (int): Zero-based end row index.

        Returns:
            ~ooodev.utils.data_type.size.Size: Size, Width is number of Columns and Height is number of Rows.

        .. versionadded:: 0.8.2
        """
        range_name = cls.get_range_str(*args, **kwargs)
        rv = mRngValues.RangeValues.from_range(range_name)
        height = rv.row_end - rv.row_start + 1
        width = rv.col_end - rv.col_start + 1
        return Size(width, height)

    # endregion get_range_size()

    # region    get_cell_str()
    @classmethod
    def _get_cell_str_addr(cls, addr: CellAddress) -> str:
        """LO Safe Method"""
        return cls._get_cell_str_col_row(col=addr.Column, row=addr.Row)

    @classmethod
    def _get_cell_str_col_row(cls, col: int, row: int) -> str:
        """LO Safe Method"""
        if col < 0 or row < 0:
            mLo.Lo.print("Cell position is negative; using A1")
            return "A1"
        return f"{cls.column_number_str(col)}{row + 1}"

    @classmethod
    def _get_cell_str_cell(cls, cell: XCell) -> str:
        """LO Safe Method"""
        return cls._get_cell_str_addr(cls._get_cell_address_cell(cell=cell))

    @overload
    @classmethod
    def get_cell_str(cls, cell_obj: mCellObj.CellObj) -> str:
        """
        Gets the cell as a string in format of ``A1``.

        |lo_safe|

        Args:
            cell_obj (CellObj): Cell object.

        Returns:
            str: Cell as str.
        """
        ...

    @overload
    @classmethod
    def get_cell_str(cls, addr: CellAddress) -> str:
        """
        Gets the cell as a string in format of ``A1``.

        |lo_safe|

        Args:
            addr (CellAddress): Cell address.

        Returns:
            str: Cell as str.
        """
        ...

    @overload
    @classmethod
    def get_cell_str(cls, cell: XCell) -> str:
        """
        Gets the cell as a string in format of ``A1``.

        |lo_safe|

        Args:
            cell (XCell): Cell.

        Returns:
            str: Cell as str.
        """
        ...

    @overload
    @classmethod
    def get_cell_str(cls, col: int, row: int) -> str:
        """
        Gets the cell as a string in format of ``A1``.

        |lo_safe|

        Args:
            col (int): Zero-based column index.
            row (int): Zero-based row index.

        Returns:
            str: Cell as str.
        """
        ...

    @classmethod
    def get_cell_str(cls, *args, **kwargs) -> str:
        """
        Gets the cell as a string in format of ``A1``.

        |lo_safe|

        Args:
            cell_obj (CellObj): Cell object.
            addr (CellAddress): Cell address.
            cell (XCell): Cell.
            col (int): Zero-based column index.
            row (int): Zero-based row index.

        Returns:
            str: Cell as str.
        """
        ordered_keys = (1, 2)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("addr", "cell", "col", "row", "cell_obj")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_cell_str() got an unexpected keyword argument")
            keys = ("addr", "cell", "col", "cell_obj")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break

            if count == 1:
                return ka

            ka[2] = kwargs.get("row", None)
            return ka

        if count not in (1, 2):
            raise TypeError("get_cell_str() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        arg1 = kargs[1]
        if count == 1:
            if isinstance(arg1, mCellObj.CellObj):
                return str(arg1)
            # def get_cell_str(addr: CellAddress) or
            # def get_cell_str(cell: XCell)
            if mInfo.Info.is_type_interface(arg1, "com.sun.star.table.XCell"):
                return cls._get_cell_str_cell(arg1)
            else:
                return cls._get_cell_str_addr(arg1)
        else:
            # def get_cell_str(col: int, row: int)
            return cls._get_cell_str_col_row(col=arg1, row=kargs[2])

    # endregion get_cell_str()

    # region get_cell_obj()
    @overload
    @classmethod
    def get_cell_obj(cls) -> mCellObj.CellObj:
        """
        Gets the cell as ``CellObj`` instance.

        |lo_unsafe|

        Returns:
            CellObj: Cell Object
        """
        ...

    @overload
    @classmethod
    def get_cell_obj(cls, cell_name: str) -> mCellObj.CellObj:
        """
        Gets the cell as ``CellObj`` instance.

        |lo_safe|

        Args:
            cell_name (str): Cell name.

        Returns:
            CellObj: Cell Object
        """
        ...

    @overload
    @classmethod
    def get_cell_obj(cls, addr: CellAddress) -> mCellObj.CellObj:
        """
        Gets the cell as ``CellObj`` instance.

        |lo_safe|

        Args:
            addr (CellAddress): Cell Address.

        Returns:
            CellObj: Cell Object
        """
        ...

    @overload
    @classmethod
    def get_cell_obj(cls, cell: XCell) -> mCellObj.CellObj:
        """
        Gets the cell as ``CellObj`` instance.

        |lo_safe|

        Args:
            cell (XCell): Cell.

        Returns:
            CellObj: Cell Object
        """
        ...

    @overload
    @classmethod
    def get_cell_obj(cls, cell_obj: mCellObj.CellObj) -> mCellObj.CellObj:
        """
        Gets the cell as ``CellObj`` instance.

        |lo_safe|

        Args:
            cell_obj (CellObj): Cell Object. If passed in the same CellObj is returned.

        Returns:
            CellObj: Cell Object
        """
        ...

    @overload
    @classmethod
    def get_cell_obj(cls, col: int, row: int) -> mCellObj.CellObj:
        """
        Gets the cell as ``CellObj`` instance.

        |lo_safe|

        Args:
            col (int): Zero-based column index.
            row (int): Zero-based row index.

        Returns:
            CellObj: Cell Object
        """
        ...

    @classmethod
    def get_cell_obj(cls, *args, **kwargs) -> mCellObj.CellObj:
        """
        Gets the cell as ``CellObj`` instance.

        |lo_safe|

        Args:
            cell_name (str): Cell name.
            addr (CellAddress): Cell Address.
            cell (XCell): Cell.
            cell_obj (CellObj): Cell Object. If passed in the same CellObj is returned.
            col (int): Zero-based column index.
            row (int): Zero-based row index.

        Returns:
            CellObj: Cell Object

        Note:
            If no args are pass in then current selected cell is returned.

        See Also:
            :py:meth:`~.calc.Calc.get_selected_cell`

        .. versionadded:: 0.8.2
        """
        kargs_len = len(kwargs)
        count = len(args) + kargs_len
        if count == 0:
            # get selected cell
            return cls.get_selected_cell()

        ordered_keys = (1, 2)

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("cell_name", "cell_obj", "addr", "cell", "col", "row")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("get_cell_obj() got an unexpected keyword argument")
            keys = ("cell_name", "cell_obj", "addr", "cell", "col")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            if count == 1:
                return ka
            ka[2] = ka.get("row")
            return ka

        def get_cell(cell_obj: mCellObj.CellObj) -> mCellObj.CellObj:

            if cell_obj.sheet_idx < 0:
                return cell_obj.set_sheet_index()
            return cell_obj

        if count not in (1, 2):
            raise TypeError("get_cell_obj() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        if count == 2:
            return get_cell(
                mCellObj.CellObj.from_idx(col_idx=kargs[1], row_idx=kargs[2], sheet_idx=cls.get_sheet_index())
            )

        arg = cast(Any, kargs[1])
        if mLo.Lo.is_uno_interfaces(arg, "com.sun.star.table.XCell"):
            return get_cell(mCellObj.CellObj.from_cell(arg.CellAddress))
        if mInfo.Info.is_instance(arg, mCellObj.CellObj):
            return get_cell(arg)

        if mInfo.Info.is_instance(arg, str):
            return get_cell(mCellObj.CellObj.from_cell(arg))

        if mInfo.Info.is_type_struct(arg, "com.sun.star.table.CellAddress"):
            return get_cell(mCellObj.CellObj.from_cell(arg))

        return mCellObj.CellObj.from_cell(cls.get_cell_str(*args, **kwargs))

    # endregion get_cell_obj()

    @staticmethod
    def column_number_str(col: int) -> str:
        """
        Creates a column Name from zero base column number.

        |lo_safe|

        Columns are numbered starting at 0 where 0 corresponds to ``A``
        They run as ``A-Z``, ``AA-AZ``, ``BA-BZ``, ..., ``IV``

        Args:
            col (int): Zero based column index

        Returns:
            str: Column Name
        """
        num = col + 1  # shift to one based.
        return mTblHelper.TableHelper.make_column_name(num)

    # endregion ------------ convert cell range address to string ------

    # region --------------- merge--------------------------------------
    # region merge_cells()
    @overload
    @classmethod
    def merge_cells(cls, cell_range: XCellRange) -> None:
        """
        Merges a range of cells.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def merge_cells(cls, cell_range: XCellRange, center: bool) -> None:
        """
        Merges a range of cells.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range.
            center (bool): Determines if the merge will be a merge and center. Default ``False``.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def merge_cells(cls, sheet: XSpreadsheet, range_name: str) -> None:
        """
        Merges a range of cells.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_name (str): Range Name such as ``A1:D5``.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def merge_cells(cls, sheet: XSpreadsheet, range_name: str, center: bool) -> None:
        """
        Merges a range of cells.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_name (str): Range Name such as ``A1:D5``.
            center (bool): Determines if the merge will be a merge and center. Default ``False``.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def merge_cells(cls, sheet: XSpreadsheet, range_obj: mRngObj.RangeObj) -> None:
        """
        Merges a range of cells.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_obj (RangeObj): Range Object.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def merge_cells(cls, sheet: XSpreadsheet, range_obj: mRngObj.RangeObj, center: bool) -> None:
        """
        Merges a range of cells.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_obj (RangeObj): Range Object.
            center (bool): Determines if the merge will be a merge and center. Default ``False``.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def merge_cells(cls, sheet: XSpreadsheet, cr_addr: CellRangeAddress) -> None:
        """
        Merges a range of cells.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            cr_addr (CellRangeAddress): Cell range Address.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def merge_cells(cls, sheet: XSpreadsheet, cr_addr: CellRangeAddress, center: bool) -> None:
        """
        Merges a range of cells.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            cr_addr (CellRangeAddress): Cell range Address.
            center (bool): Determines if the merge will be a merge and center. Default ``False``.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def merge_cells(cls, col_start: int, row_start: int, col_end: int, row_end: int, center: bool) -> None:
        """
        Merges a range of cells.

        |lo_safe|

        Args:
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.
            center (bool): Determines if the merge will be a merge and center. Default ``False``.

        Returns:
            None:
        """
        ...

    @classmethod
    def merge_cells(cls, *args, **kwargs) -> None:
        """
        Merges a range of cells.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            center (bool): Determines if the merge will be a merge and center. Default ``False``.
            range_name (str): Range Name such as ``A1:D5``.
            range_obj (RangeObj): Range Object.
            cr_addr (CellRangeAddress): Cell range Address.
            cell_range (XCellRange): Cell Range.
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.

        Returns:
            None:

        See Also:
            - :py:meth:`.Calc.unmerge_cells`
            - :py:meth:`.Calc.is_merged_cells`

        .. versionadded:: 0.8.4
        """
        # center must be removed from args if it exist so the rest of the args can be passed to get_cell_range()
        center = None
        kw = kwargs.copy()
        lst_args = list(args)
        args_len = len(lst_args)
        if "center" in kw:
            center = bool(kw["center"])
            del kw["center"]

        if center is None and args_len > 0 and isinstance(lst_args[-1], bool):
            center = lst_args.pop()

        cell_range = cls.get_cell_range(*lst_args, **kw)
        x_merge = mLo.Lo.qi(XMergeable, cell_range, True)
        x_merge.merge(True)
        if center:
            mProps.Props.set(cell_range, HoriJustify=CellHoriJustify.CENTER, VertJustify=CellVertJustify2.CENTER)

    # endregion merge_cells()

    # region unmerge_cells()
    @overload
    @classmethod
    def unmerge_cells(cls, cell_range: XCellRange) -> None:
        """
        Removes merging from a range of cells.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def unmerge_cells(cls, sheet: XSpreadsheet, range_name: str) -> None:
        """
        Removes merging from a range of cells.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_name (str): Range Name such as ``A1:D5``.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def unmerge_cells(cls, sheet: XSpreadsheet, range_obj: mRngObj.RangeObj) -> None:
        """
        Removes merging from a range of cells.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_obj (RangeObj): Range Object.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def unmerge_cells(cls, sheet: XSpreadsheet, cr_addr: CellRangeAddress) -> None:
        """
        Removes merging from a range of cells.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            cr_addr (CellRangeAddress): Cell range Address.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def unmerge_cells(cls, col_start: int, row_start: int, col_end: int, row_end: int) -> None:
        """
        Removes merging from a range of cells.

        |lo_safe|

        Args:
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.

        Returns:
            None:
        """
        ...

    @classmethod
    def unmerge_cells(cls, *args, **kwargs) -> None:
        """
        Removes merging from a range of cells.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_name (str): Range Name such as ``A1:D5``.
            range_obj (RangeObj): Range Object.
            cr_addr (CellRangeAddress): Cell range Address.
            cell_range (XCellRange): Cell Range.
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.

        Returns:
            None:

        See Also:
            - :py:meth:`.Calc.merge_cells`
            - :py:meth:`.Calc.is_merged_cells`

        .. versionadded:: 0.8.4
        """
        cell_range = cls.get_cell_range(*args, **kwargs)
        x_merge = mLo.Lo.qi(XMergeable, cell_range, True)
        x_merge.merge(False)
        # XMergeable

    # endregion unmerge_cells()

    # region is_merged_cells()
    @overload
    @classmethod
    def is_merged_cells(cls, cell_range: XCellRange) -> bool:
        """
        Gets is a range of cells is merged.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range.

        Returns:
            bool: ``True`` if range is merged; Otherwise, ``False``
        """
        ...

    @overload
    @classmethod
    def is_merged_cells(cls, sheet: XSpreadsheet, range_name: str) -> bool:
        """
        Gets is a range of cells is merged.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_name (str): Range Name such as ``A1:D5``.

        Returns:
            bool: ``True`` if range is merged; Otherwise, ``False``
        """
        ...

    @overload
    @classmethod
    def is_merged_cells(cls, sheet: XSpreadsheet, range_obj: mRngObj.RangeObj) -> bool:
        """
        Gets is a range of cells is merged.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_obj (RangeObj): Range Object.

        Returns:
            bool: ``True`` if range is merged; Otherwise, ``False``
        """
        ...

    @overload
    @classmethod
    def is_merged_cells(cls, sheet: XSpreadsheet, cr_addr: CellRangeAddress) -> bool:
        """
        Gets is a range of cells is merged.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            cr_addr (CellRangeAddress): Cell range Address.

        Returns:
            bool: ``True`` if range is merged; Otherwise, ``False``
        """
        ...

    @overload
    @classmethod
    def is_merged_cells(cls, col_start: int, row_start: int, col_end: int, row_end: int) -> bool:
        """
        Gets is a range of cells is merged.

        |lo_safe|

        Args:
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.

        Returns:
            bool: ``True`` if range is merged; Otherwise, ``False``
        """
        ...

    @classmethod
    def is_merged_cells(cls, *args, **kwargs) -> bool:
        """
        Gets is a range of cells is merged.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_name (str): Range Name such as ``A1:D5``.
            range_obj (RangeObj): Range Object.
            cr_addr (CellRangeAddress): Cell range Address.
            cell_range (XCellRange): Cell Range.
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.

        Returns:
            bool: ``True`` if range is merged; Otherwise, ``False``

        See Also:
            - :py:meth:`.Calc.merge_cells`
            - :py:meth:`.Calc.unmerge_cells`

        .. versionadded:: 0.8.4
        """
        cell_range = cls.get_cell_range(*args, **kwargs)
        x_merge = mLo.Lo.qi(XMergeable, cell_range, True)
        return x_merge.getIsMerged()

    # endregion is_merged_cells()

    # endregion ------------ merge--------------------------------------

    # region --------------- search ------------------------------------

    @staticmethod
    def find_all(srch: XSearchable, sd: XSearchDescriptor) -> List[XCellRange] | None:
        """
        Searches spreadsheet and returns a list of Cell Ranges that match search criteria.

        |lo_safe|

        Args:
            srch (XSearchable): Searchable object.
            sd (XSearchDescriptor): Search description.

        Returns:
            List[XCellRange] | None: A list of cell ranges on success; Otherwise, ``None``.

        .. collapse:: Example

            .. code-block:: python

                from ooodev.loader.lo import Lo
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

        See Also:
            `LibreOffice API SearchDescriptor <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1util_1_1SearchDescriptor.html>`_
        """
        con = srch.findAll(sd)
        if con is None:
            mLo.Lo.print("Match result is null")
            return None
        c_count = con.getCount()
        if c_count == 0:
            mLo.Lo.print("No matches found")
            return None

        crs = []
        for i in range(c_count):
            try:
                cr = mLo.Lo.qi(XCellRange, con.getByIndex(i))
                if cr is None:
                    continue
                crs.append(cr)
            except Exception:
                mLo.Lo.print(f"Could not access match index {i}")
        if not crs:
            mLo.Lo.print(f"Found {c_count} matches but unable to access any match")
            return None
        return crs

    @staticmethod
    def get_sheet_name_from_code_name(doc: XSpreadsheetDocument, code_name: str) -> str:
        """
        Gets the sheet name from the code name.

        Args:
            doc (XSpreadsheetDocument): Spreadsheet document.
            code_name (str): The sheet code name. Case insensitive.

        Returns:
            str: The sheet name or an empty string if not found.

        .. versionadded:: 0.44.1
        """
        if not code_name:
            return ""
        s = code_name.casefold()
        result = ""
        sheets = doc.getSheets()
        for sheet in sheets:  # type: ignore
            if sheet.CodeName.casefold() == s:  # type: ignore
                result = sheet.Name  # type: ignore
                break
        return result

    # endregion ------------ search ------------------------------------

    # region --------------- cell decoration ---------------------------

    @staticmethod
    def create_cell_style(doc: XSpreadsheetDocument, style_name: str) -> XStyle:
        """
        Creates a style.

        |lo_unsafe|

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document.
            style_name (str): Style name.

        Raises:
            Exception: if unable to create style.
            MissingInterfaceError: if unable to obtain interface.

        Returns:
            XStyle: Newly created style.
        """
        comp_doc = mLo.Lo.qi(XComponent, doc, raise_err=True)
        style_families = mInfo.Info.get_style_container(doc=comp_doc, family_style_name="CellStyles")
        style = mLo.Lo.create_instance_msf(XStyle, "com.sun.star.style.CellStyle", raise_err=True)
        #   "com.sun.star.sheet.TableCellStyle"  result in style == None ??
        try:
            style_families.insertByName(style_name, style)
            return style
        except Exception as e:
            raise Exception(f"Unable to create style: {style_name}") from e

    # region    change_style()

    @overload
    @classmethod
    def change_style(cls, sheet: XSpreadsheet, style_name: str, cell_range: XCellRange) -> bool:
        """
        Changes style of a range of cells.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            style_name (str): Name of style to apply.
            cell_range (XCellRange): Cell range to apply style to.

        Returns:
            bool: ``True`` if style has been changed; Otherwise, ``False``.
        """
        ...

    @overload
    @classmethod
    def change_style(cls, sheet: XSpreadsheet, style_name: str, range_name: str) -> bool:
        """
        Changes style of a range of cells.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            style_name (str): Name of style to apply.
            range_name (str): Range to apply style to such as ``A1:E23``.

        Returns:
            bool: ``True`` if style has been changed; Otherwise, ``False``.
        """
        ...

    @overload
    @classmethod
    def change_style(cls, sheet: XSpreadsheet, style_name: str, range_obj: mRngObj.RangeObj) -> bool:
        """
        Changes style of a range of cells.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            style_name (str): Name of style to apply.
            range_obj (RangeObj): Range Object.

        Returns:
            bool: ``True`` if style has been changed; Otherwise, ``False``.
        """
        ...

    @overload
    @classmethod
    def change_style(
        cls, sheet: XSpreadsheet, style_name: str, start_col: int, start_row: int, end_col: int, end_row: int
    ) -> bool:
        """
        Changes style of a range of cells.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            style_name (str): Name of style to apply.
            start_col (int): Zero-base start column index.
            start_row (int): Zero-base start row index.
            end_col (int): Zero-base end column index.
            end_row (int): Zero-base end row index.

        Returns:
            bool: ``True`` if style has been changed; Otherwise, ``False``.
        """
        ...

    @classmethod
    def change_style(cls, *args, **kwargs) -> bool:
        """
        Changes style of a range of cells.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            style_name (str): Name of style to apply.
            cell_range (XCellRange): Cell range to apply style to.
            range_name (str): Range to apply style to such as ``A1:E23``.
            range_obj (RangeObj): Range Object.
            start_col (int): Zero-base start column index.
            start_row (int): Zero-base start row index.
            end_col (int): Zero-base end column index.
            end_row (int): Zero-base end row index.

        Returns:
            bool: ``True`` if style has been changed; Otherwise, ``False``.
        """
        ordered_keys = (1, 2, 3, 4, 5, 6)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = (
                "sheet",
                "style_name",
                "range_name",
                "range_obj",
                "cell_range",
                "start_col",
                "start_row",
                "end_col",
                "end_row",
            )
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("change_style() got an unexpected keyword argument")
            ka[1] = kwargs.get("sheet", None)
            ka[2] = kwargs.get("style_name", None)
            keys = ("range_name", "range_obj", "start_col", "cell_range")
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

        if count not in (3, 6):
            raise TypeError("change_style() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        sheet = cast(XSpreadsheet, kargs[1])
        style_name = cast(str, kargs[2])
        if count == 3:
            arg3 = kargs[3]
            if isinstance(arg3, str):
                # change_style(sheet: XSpreadsheet, style_name: str, range_name: str)

                cell_range = cls._get_cell_range_rng_name(
                    sheet=sheet, range_name=cls.get_safe_rng_str(arg3)
                )  # 1 sheet, 3 range_name
                if cell_range is None:
                    return False
            elif isinstance(arg3, mRngObj.RangeObj):
                cell_range = cls.get_cell_range(sheet=sheet, range_obj=arg3)
                if cell_range is None:
                    return False
            else:
                cell_range = arg3
            mProps.Props.set(cell_range, CellStyle=style_name)  # 2 style_name
            return style_name == mProps.Props.get(cell_range, "CellStyle")
        else:
            # def change_style(sheet: XSpreadsheet, style_name: str, x1: int, y1: int, x2: int, y2:int)
            cell_range = cls._get_cell_range_col_row(
                sheet=sheet, start_col=kargs[3], start_row=kargs[4], end_col=kargs[5], end_row=kargs[6]
            )
            mProps.Props.set(cell_range, CellStyle=style_name)  # 2 style_name
            return style_name == mProps.Props.get(cell_range, "CellStyle")

        # endregion change_style()

    # region    add_border()
    @classmethod
    def _add_border_sht_rng(cls, cargs: CellCancelArgs) -> None:
        """Lo Safe Method"""
        cargs.event_data["color"] = CommonColor.BLACK
        cls._add_border_sht_rng_color(cargs)  # color black

    @classmethod
    def _add_border_sht_rng_color(cls, cargs: CellCancelArgs) -> None:
        """Lo Safe Method"""
        vals = (
            cls.BorderEnum.LEFT_BORDER
            | cls.BorderEnum.RIGHT_BORDER
            | cls.BorderEnum.TOP_BORDER
            | cls.BorderEnum.BOTTOM_BORDER
        )
        cargs.event_data["border_vals"] = vals
        cls._add_border_sht_rng_color_vals(cargs)

    @classmethod
    def _add_border_sht_rng_color_vals(
        cls,
        cargs: CellCancelArgs,
    ) -> None:
        """Lo Safe Method"""
        _Events().trigger(CalcNamedEvent.CELLS_BORDER_ADDING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)
        cell_range = cast(XCellRange, cargs.cells)
        color = int(cargs.event_data["color"])
        bvs = cls.BorderEnum(int(cargs.event_data["border_vals"]))
        line = BorderLine2()  # create the border line
        border = cast(TableBorder2, mProps.Props.get(cell_range, "TableBorder2"))
        inner_line = cast(BorderLine2, mProps.Props.get(cell_range, "TopBorder2"))

        line.Color = color  # type: ignore
        line.InnerLineWidth = 0
        line.LineDistance = 0
        line.OuterLineWidth = 100

        # inner_line = BorderLine2()  # create the border line
        # inner_line.Color = 0
        # inner_line.LineWidth = 0
        # inner_line.InnerLineWidth = 0
        # inner_line.LineDistance = 0
        # inner_line.LineStyle = 0
        # inner_line.OuterLineWidth = 0
        # border = TableBorder2()

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
        mProps.Props.set(
            cell_range,
            TopBorder2=inner_line,
            RightBorder2=inner_line,
            BottomBorder2=inner_line,
            LeftBorder2=inner_line,
            TableBorder2=border,
        )

    @overload
    @classmethod
    def add_border(cls, sheet: XSpreadsheet, cell_range: XCellRange) -> XCellRange:
        """
        Adds borders to a cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_range (XCellRange): Cell range.

        Returns:
            XCellRange: Range borders that are affected.
        """
        ...

    @overload
    @classmethod
    def add_border(cls, sheet: XSpreadsheet, range_name: str) -> XCellRange:
        """
        Adds borders to a cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range Name such as ``A1:F9``.

        Returns:
            XCellRange: Range borders that are affected.
        """
        ...

    @overload
    @classmethod
    def add_border(cls, sheet: XSpreadsheet, cell_range: XCellRange, color: Color) -> XCellRange:
        """
        Adds borders to a cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_range (XCellRange): Cell range.
            color (Color): RGB color.

        Returns:
            XCellRange: Range borders that are affected.
        """
        ...

    @overload
    @classmethod
    def add_border(cls, sheet: XSpreadsheet, range_name: str, color: Color) -> XCellRange:
        """
        Adds borders to a cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range Name such as ``A1:F9``.
            color (Color): RGB color

        Returns:
            XCellRange: Range borders that are affected.
        """
        ...

    @overload
    @classmethod
    def add_border(
        cls, sheet: XSpreadsheet, cell_range: XCellRange, color: Color, border_vals: BorderEnum
    ) -> XCellRange:
        """
        Adds borders to a cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_range (XCellRange): Cell range.
            color (Color): RGB color.
            border_vals (BorderEnum): Determines what borders are applied.

        Returns:
            XCellRange: Range borders that are affected.
        """
        ...

    @overload
    @classmethod
    def add_border(cls, sheet: XSpreadsheet, range_name: str, color: Color, border_vals: BorderEnum) -> XCellRange:
        """
        Adds borders to a cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range Name such as ``A1:F9``.
            color (Color): RGB color
            border_vals (BorderEnum): Determines what borders are applied.

        Returns:
            XCellRange: Range borders that are affected.
        """
        ...

    @classmethod
    def add_border(cls, *args, **kwargs) -> XCellRange:
        """
        Adds borders to a cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_range (XCellRange): Cell range.
            range_name (str): Range Name such as ``A1:F9``.
            color (Color): RGB color
            border_vals (BorderEnum): Determines what borders are applied.

        Raises:
            CancelEventError: If CELLS_BORDER_ADDING event is canceled.

        Returns:
            XCellRange: Range borders that are affected.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_BORDER_ADDING` :eventref:`src-docs-cell-event-border-adding`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_BORDER_ADDED` :eventref:`src-docs-cell-event-border-added`

        Note:
            Event args ``cells`` is of type ``XCellRange``.

            Event args ``event_data`` is a dictionary containing ``color`` and ``border_vals``.

        See Also:
            - :py:meth:`~.calc.Calc.remove_border`
            - :py:meth:`~.calc.Calc.highlight_range`
            - :ref:`ch22_adding_borders`
        """
        ordered_keys = (1, 2, 3, 4)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("sheet", "range_name", "cell_range", "color", "border_vals")
            check = all(key in valid_keys for key in kwargs)
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

        if count not in (2, 3, 4):
            raise TypeError("add_border() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        sheet = cast(XSpreadsheet, kargs[1])
        if isinstance(kargs[2], str):
            cell_range = sheet.getCellRangeByName(kargs[2])
        else:
            cell_range = cast(XCellRange, kargs[2])

        cargs = CellCancelArgs(cls)
        cargs.sheet = sheet
        cargs.cells = cell_range
        cargs.event_data = {}
        if count == 2:
            # add_border(sheet: XSpreadsheet, cell_range: str)
            # add_border(sheet: XSpreadsheet, range_name: str)
            cls._add_border_sht_rng(cargs)
        elif count == 3:
            # add_border(sheet: XSpreadsheet, cell_range: XCellRange, color: int)
            #  add_border(sheet: XSpreadsheet, range_name: str, color: int)
            cargs.event_data["color"] = kargs[3]
            cls._add_border_sht_rng_color(cargs)
        else:
            # add_border(sheet: XSpreadsheet, cell_range: XCellRange, color: int, border_vals: int)
            # add_border(sheet: XSpreadsheet, range_name: str, color: int, border_vals: int)
            # add_border(sheet: XSpreadsheet, cell_range: XCellRange, color: int, border_vals: Calc.BorderEnum)
            # add_border(sheet: XSpreadsheet, range_name: str, color: int, border_vals: BorderEnum)
            cargs.event_data["color"] = kargs[3]
            cargs.event_data["border_vals"] = kargs[4]

            cls._add_border_sht_rng_color_vals(cargs)
        _Events().trigger(CalcNamedEvent.CELLS_BORDER_ADDED, CellArgs.from_args(cargs))
        return cell_range

    # endregion add_border()

    # region    remove_border()
    @classmethod
    def _remove_border_sht_rng(cls, cargs: CellCancelArgs) -> None:
        """LO Safe Method"""
        vals = (
            cls.BorderEnum.LEFT_BORDER
            | cls.BorderEnum.RIGHT_BORDER
            | cls.BorderEnum.TOP_BORDER
            | cls.BorderEnum.BOTTOM_BORDER
        )
        cargs.event_data = vals
        cls._remove_border_sht_rng_vals(cargs)

    @classmethod
    def _remove_border_sht_rng_vals(
        cls,
        cargs: CellCancelArgs,
    ) -> None:
        """LO Safe Method"""
        _Events().trigger(CalcNamedEvent.CELLS_BORDER_REMOVING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)
        border_vals = cast(Calc.BorderEnum, cargs.event_data)
        cell_range = cast(XCellRange, cargs.cells)
        line = BorderLine2()  # create the border line
        border = cast(TableBorder2, mProps.Props.get(cell_range, "TableBorder2"))
        line = cast(BorderLine2, mProps.Props.get(cell_range, "TopBorder2"))
        inner_line = cast(BorderLine2, mProps.Props.get(cell_range, "TopBorder2"))
        line.Color = 0  # type: ignore
        line.LineWidth = 0
        line.InnerLineWidth = 0
        line.LineDistance = 0
        line.LineStyle = 0
        line.OuterLineWidth = 0

        bvs = cls.BorderEnum(int(border_vals))
        # border = TableBorder2()

        if (bvs & cls.BorderEnum.TOP_BORDER) == cls.BorderEnum.TOP_BORDER:
            border.TopLine = line
            border.IsTopLineValid = False

        if (bvs & cls.BorderEnum.BOTTOM_BORDER) == cls.BorderEnum.BOTTOM_BORDER:
            border.BottomLine = line
            border.IsBottomLineValid = False

        if (bvs & cls.BorderEnum.LEFT_BORDER) == cls.BorderEnum.LEFT_BORDER:
            border.LeftLine = line
            border.IsLeftLineValid = False

        if (bvs & cls.BorderEnum.RIGHT_BORDER) == cls.BorderEnum.RIGHT_BORDER:
            border.RightLine = line
            border.IsRightLineValid = False

        mProps.Props.set(
            cell_range,
            TableBorder2=border,
            TopBorder2=inner_line,
            RightBorder2=inner_line,
            BottomBorder2=inner_line,
            LeftBorder2=inner_line,
        )

    @overload
    @classmethod
    def remove_border(cls, sheet: XSpreadsheet, cell_range: XCellRange) -> XCellRange:
        """
        Removes borders of a cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_range (XCellRange): Cell range.

        Returns:
            XCellRange: Range borders that are affected.
        """
        ...

    @overload
    @classmethod
    def remove_border(cls, sheet: XSpreadsheet, range_name: str) -> XCellRange:
        """
        Removes borders of a cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range Name such as ``A1:F9``.

        Returns:
            XCellRange: Range borders that are affected.
        """
        ...

    @overload
    @classmethod
    def remove_border(cls, sheet: XSpreadsheet, range_obj: mRngObj.RangeObj) -> XCellRange:
        """
        Removes borders of a cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet
            range_obj (RangeObj): Range Object.

        Returns:
            XCellRange: Range borders that are affected.
        """
        ...

    @overload
    @classmethod
    def remove_border(cls, sheet: XSpreadsheet, cell_range: XCellRange, border_vals: BorderEnum) -> XCellRange:
        """
        Removes borders of a cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_range (XCellRange): Cell range.
            border_vals (BorderEnum): Determines what borders are applied.

        Returns:
            XCellRange: Range borders that are affected.
        """
        ...

    @overload
    @classmethod
    def remove_border(cls, sheet: XSpreadsheet, range_name: str, border_vals: BorderEnum) -> XCellRange:
        """
        Removes borders of a cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range Name such as ``A1:F9``.
            border_vals (BorderEnum): Determines what borders are applied.

        Returns:
            XCellRange: Range borders that are affected.
        """
        ...

    @overload
    @classmethod
    def remove_border(cls, sheet: XSpreadsheet, range_obj: mRngObj.RangeObj, border_vals: BorderEnum) -> XCellRange:
        """
        Removes borders of a cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_obj (RangeObj): Range Object.
            border_vals (BorderEnum): Determines what borders are applied.

        Returns:
            XCellRange: Range borders that are affected.
        """
        ...

    @classmethod
    def remove_border(cls, *args, **kwargs) -> XCellRange:
        """
        Removes borders of a cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_range (XCellRange): Cell range.
            range_name (str): Range Name such as ``A1:F9``.
            range_obj (RangeObj): Range Object.
            border_vals (BorderEnum): Determines what borders are applied.

        Raises:
            CancelEventError: If CELLS_BORDER_REMOVING event is canceled.

        Returns:
            XCellRange: Range borders that are affected.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_BORDER_REMOVING` :eventref:`src-docs-cell-event-border-removing`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_BORDER_REMOVED` :eventref:`src-docs-cell-event-border-removed`

        Note:
            Event args ``cells`` is of type ``XCellRange``.

            Event arg properties modified on CELLS_BORDER_REMOVING it is reflected in this method.

        See Also:
            :py:meth:`~.calc.Calc.add_border`
        """
        ordered_keys = (1, 2, 3)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("sheet", "range_name", "range_obj", "cell_range", "border_vals")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("remove_border() got an unexpected keyword argument")
            ka[1] = kwargs.get("sheet", None)
            keys = ("range_name", "range_obj", "cell_range")
            for key in keys:
                if key in kwargs:
                    ka[2] = kwargs[key]
                    break
            if count == 2:
                return ka
            ka[3] = kwargs.get("border_vals", None)
            return ka

        if count not in (2, 3):
            raise TypeError("remove_border() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        sheet = cast(XSpreadsheet, kargs[1])
        if isinstance(kargs[2], (str, mRngObj.RangeObj)):
            cell_range = sheet.getCellRangeByName(str(kargs[2]))
        else:
            cell_range = cast(XCellRange, kargs[2])

        cargs = CellCancelArgs(cls)
        cargs.sheet = sheet
        cargs.cells = cell_range

        if count == 2:
            # remove_border(cls, sheet: XSpreadsheet, cell_range: XCellRange)
            # remove_border(cls, sheet: XSpreadsheet, range_name: str)
            cls._remove_border_sht_rng(cargs)
        else:
            # remove_border(cls, sheet: XSpreadsheet, cell_range: XCellRange, border_vals: BorderEnum)
            # remove_border(cls, sheet: XSpreadsheet, range_name: str, border_vals: BorderEnum)
            cargs.event_data = kargs[3]
            cls._remove_border_sht_rng_vals(cargs)

        _Events().trigger(CalcNamedEvent.CELLS_BORDER_REMOVED, CellArgs.from_args(cargs))
        return cell_range

    # endregion remove_border()

    # region    highlight_range()
    @overload
    @classmethod
    def highlight_range(cls, sheet: XSpreadsheet, headline: str, cell_range: XCellRange) -> XCell:
        """
        Draw a colored border around the range and write a headline in the
        top-left cell of the range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            headline (str): Headline.
            cell_range (XCellRange): Cell Range.

        Returns:
            XCell: First cell of range that headline ia applied on.
        """
        ...

    @overload
    @classmethod
    def highlight_range(cls, sheet: XSpreadsheet, headline: str, cell_range: XCellRange, color: Color) -> XCell:
        """
        Draw a colored border around the range and write a headline in the
        top-left cell of the range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            headline (str): Headline.
            cell_range (XCellRange): Cell Range.
            color (~ooodev.utils.color.Color): RGB color.

        Returns:
            XCell: First cell of range that headline ia applied on.
        """
        ...

    @overload
    @classmethod
    def highlight_range(cls, sheet: XSpreadsheet, headline: str, range_name: str) -> XCell:
        """
        Draw a colored border around the range and write a headline in the
        top-left cell of the range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            headline (str): Headline.
            range_name (str): Range Name such as ``A1:F9``.

        Returns:
            XCell: First cell of range that headline ia applied on.
        """
        ...

    @overload
    @classmethod
    def highlight_range(cls, sheet: XSpreadsheet, headline: str, range_obj: mRngObj.RangeObj) -> XCell:
        """
        Draw a colored border around the range and write a headline in the
        top-left cell of the range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            headline (str): Headline.
            range_obj (RangeObj): Range Object
            color (~ooodev.utils.color.Color): RGB color.

        Returns:
            XCell: First cell of range that headline ia applied on.
        """
        ...

    @overload
    @classmethod
    def highlight_range(cls, sheet: XSpreadsheet, headline: str, range_name: str, color: Color) -> XCell:
        """
        Draw a colored border around the range and write a headline in the
        top-left cell of the range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            headline (str): Headline.
            range_name (str): Range Name such as 'A1:F9'.
            color (~ooodev.utils.color.Color): RGB color.

        Returns:
            XCell: First cell of range that headline ia applied on.
        """
        ...

    @overload
    @classmethod
    def highlight_range(cls, sheet: XSpreadsheet, headline: str, range_obj: mRngObj.RangeObj, color: Color) -> XCell:
        """
        Draw a colored border around the range and write a headline in the
        top-left cell of the range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            headline (str): Headline.
            cell_range (XCellRange): Cell Range.
            range_obj (RangeObj): Range Object
            color (~ooodev.utils.color.Color): RGB color.

        Returns:
            XCell: First cell of range that headline ia applied on.
        """
        ...

    @classmethod
    def highlight_range(cls, *args, **kwargs) -> XCell | None:
        """
        Draw a colored border around the range and write a headline in the
        top-left cell of the range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            headline (str): Headline.
            cell_range (XCellRange): Cell Range.
            range_name (str): Range Name such as 'A1:F9'.
            range_obj (RangeObj): Range Object
            color (~ooodev.utils.color.Color): RGB color.

        Raises:
            CancelEventError: If CELLS_HIGH_LIGHTING event is canceled.

        Returns:
            XCell: First cell of range that headline ia applied on.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_HIGH_LIGHTING` :eventref:`src-docs-cell-event-highlighting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_HIGH_LIGHTED` :eventref:`src-docs-cell-event-highlighted`

        Note:
            Event args ``cells`` is of type ``XCellRange``.

            Event args ``event_data`` is a dictionary containing ``color`` and ``headline``.

        See Also:
            - :py:meth:`~.calc.Calc.add_border`
            - :ref:`ch24_creating_border_headline`
        """
        ordered_keys = (1, 2, 3, 4)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("sheet", "headline", "range_name", "range_obj", "cell_range", "color")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("highlight_range() got an unexpected keyword argument")
            ka[1] = kwargs.get("sheet", None)
            ka[2] = kwargs.get("headline", None)
            keys = ("range_name", "range_obj", "cell_range")
            for key in keys:
                if key in kwargs:
                    ka[3] = kwargs[key]
                    break
            if count == 3:
                return ka
            ka[4] = kwargs.get("color", None)
            return ka

        if count not in (3, 4):
            raise TypeError("highlight_range() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        sheet = cast(XSpreadsheet, kargs[1])
        arg3 = kargs[3]
        if isinstance(arg3, str):
            cell_range = sheet.getCellRangeByName(arg3)
        elif isinstance(arg3, mRngObj.RangeObj):
            cell_range = cls.get_cell_range(sheet, arg3)
        else:
            cell_range = cast(XCellRange, arg3)

        color = CommonColor.LIGHT_BLUE if count == 3 else cast(Color, kargs[4])
        cargs = CellCancelArgs(Calc.highlight_range.__qualname__)
        cargs.cells = cell_range
        cargs.sheet = sheet
        cargs.event_data = {"color": color, "headline": kargs[2]}
        _Events().trigger(CalcNamedEvent.CELLS_HIGH_LIGHTING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)
        cls._add_border_sht_rng_color(cargs)

        # color the headline
        addr = cls._get_address_cell(cell_range=cargs.cells)
        header_range = cls._get_cell_range_col_row(
            sheet=cargs.sheet,
            start_col=addr.StartColumn,
            start_row=addr.StartRow,
            end_col=addr.EndColumn,
            end_row=addr.StartRow,
        )
        first_cell = cls._get_cell_cell_rng(cell_range=header_range, col=0, row=0)
        cls._set_val_by_cell(value=cargs.event_data["headline"], cell=first_cell)
        _Events().trigger(CalcNamedEvent.CELLS_HIGH_LIGHTED, CellArgs.from_args(cargs))
        return first_cell

    # endregion highlight_range()

    @classmethod
    def set_col_width(cls, sheet: XSpreadsheet, width: int | UnitT, idx: int) -> XCellRange | None:
        """
        Sets column width. width is in ``mm``, e.g. ``6``

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet
            width (int, UnitT): Width in ``mm`` units or :ref:`proto_unit_obj`.
            idx (int): Index of column.

        Raises:
            CancelEventError: If ``SHEET_COL_WIDTH_SETTING`` event is canceled.

        Returns:
            XCellRange | None: Column cell range that width is applied on or ``None`` if column width <= 0

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_COL_WIDTH_SETTING` :eventref:`src-docs-sheet-event-col-width-setting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_COL_WIDTH_SET` :eventref:`src-docs-sheet-event-col-width-set`

        Note:
            Event args ``index`` is set to ``idx`` value, ``event_data`` is set to ``width`` value (``mm100`` units).

        .. versionchanged:: 0.9.4
            width can now also be ``UnitT``
        """
        try:
            col_width = width.get_value_mm100()  # type: ignore
        except AttributeError:
            col_width = UnitConvert.convert_mm_mm100(width)  # type: ignore
        cargs = SheetCancelArgs(Calc.set_col_width.__qualname__)
        cargs.sheet = sheet
        cargs.index = idx
        cargs.event_data = col_width
        _Events().trigger(CalcNamedEvent.SHEET_COL_WIDTH_SETTING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)
        col_width = cargs.event_data
        if col_width <= 0:
            mLo.Lo.print("Width must be greater then 0")
            return None
        cell_range = cls.get_col_range(sheet=cargs.sheet, idx=cargs.index)
        mProps.Props.set(cell_range, Width=col_width)
        _Events().trigger(CalcNamedEvent.SHEET_COL_WIDTH_SET, SheetArgs.from_args(cargs))
        return cell_range

    @classmethod
    def set_row_height(
        cls,
        sheet: XSpreadsheet,
        height: int | UnitT,
        idx: int,
    ) -> XCellRange | None:
        """
        Sets column width. height is in ``mm``, e.g. ``6``.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet
            height (int, UnitT): Width in ``mm`` units or :ref:`proto_unit_obj`.
            idx (int): Index of Row

        Raises:
            CancelEventError: If ``SHEET_ROW_HEIGHT_SETTING`` event is canceled.

        Returns:
            XCellRange | None: Row cell range that height is applied on or None if height <= ``0``.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_ROW_HEIGHT_SETTING` :eventref:`src-docs-sheet-event-row-height-setting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_ROW_HEIGHT_SET` :eventref:`src-docs-sheet-event-row-height-set`

        Note:
            Event args ``index`` is set to ``idx`` value, ``event_data`` is set to ``height`` value (``mm100`` units).

        .. versionchanged:: 0.9.4
            width can now also be ``UnitT``
        """
        try:
            row_height = height.get_value_mm100()  # type: ignore
        except AttributeError:
            row_height = UnitConvert.convert_mm_mm100(height)  # type: ignore
        cargs = SheetCancelArgs(Calc.set_row_height.__qualname__)
        cargs.sheet = sheet
        cargs.index = idx
        cargs.event_data = row_height
        _Events().trigger(CalcNamedEvent.SHEET_ROW_HEIGHT_SETTING, cargs)
        if cargs.cancel:
            raise mEx.CancelEventError(cargs)
        idx = cargs.index
        row_height = cargs.event_data
        if row_height <= 0:
            mLo.Lo.print("Height must be greater then 0")
            return None
        cell_range = cls.get_row_range(sheet=cargs.sheet, idx=idx)
        # mInfo.Info.show_services(obj_name="Cell range for a row", obj=cell_range)
        mProps.Props.set(cell_range, Height=row_height)
        _Events().trigger(CalcNamedEvent.SHEET_ROW_HEIGHT_SET, SheetArgs.from_args(cargs))
        return cell_range

    # endregion ------------ cell decoration ---------------------------

    # region --------------- scenarios ---------------------------------

    @staticmethod
    def insert_scenario(
        sheet: XSpreadsheet, range_name: str | mRngObj.RangeObj, vals: Table, name: str, comment: str
    ) -> XScenario:
        """
        Insert a scenario into sheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str | RangeObj): Range name.
            vals (Table): 2d array of values.
            name (str): Scenario name.
            comment (str): Scenario description.

        Raises:
            MissingInterfaceError: If a required interface is missing.

        Returns:
            XScenario: the newly created scenario.

        Note:
            A LibreOffice Calc scenario is a set of cell values that can be used within your calculations.
            You assign a name to every scenario on your sheet. Define several scenarios on the same sheet,
            each with some different values in the cells. Then you can easily switch the sets of cell values
            by their name and immediately observe the results. Scenarios are a tool to test out "what-if" questions.

        See Also:
            `Using Scenarios <https://help.libreoffice.org/latest/en-US/text/scalc/guide/scenario.html>`_
        """
        # get the cell range with the given address
        cell_range = sheet.getCellRangeByName(str(range_name))

        # create the range address sequence
        addr = mLo.Lo.qi(XCellRangeAddressable, cell_range, raise_err=True)
        cr_addr = (addr.getRangeAddress(),)

        # create the scenario
        supp = mLo.Lo.qi(XScenariosSupplier, sheet, raise_err=True)
        scenarios = supp.getScenarios()
        scenarios.addNewByName(name, cr_addr, comment)

        # insert the values into the range
        cr_data = mLo.Lo.qi(XCellRangeData, cell_range, raise_err=True)
        cr_data.setDataArray(vals)  # type: ignore

        supp = mLo.Lo.qi(XScenariosSupplier, sheet, raise_err=True)
        scenarios = supp.getScenarios()
        return mLo.Lo.qi(XScenario, scenarios.getByName(name), raise_err=True)

    @staticmethod
    def apply_scenario(sheet: XSpreadsheet, name: str) -> XScenario:
        """
        Applies scenario.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            name (str): Scenario name to apply.

        Raises:
            Exception: If scenario is not able to be applied.

        Returns:
            XScenario: the applied scenario.

        Note:
            A LibreOffice Calc scenario is a set of cell values that can be used within your calculations.
            You assign a name to every scenario on your sheet. Define several scenarios on the same sheet,
            each with some different values in the cells. Then you can easily switch the sets of cell values
            by their name and immediately observe the results. Scenarios are a tool to test out "what-if" questions.

        See Also:
            `Using Scenarios <https://help.libreoffice.org/latest/en-US/text/scalc/guide/scenario.html>`_
        """
        try:
            # get the scenario set
            supp = mLo.Lo.qi(XScenariosSupplier, sheet, raise_err=True)
            scenarios = supp.getScenarios()

            # get the scenario and activate it
            scenario = mLo.Lo.qi(XScenario, scenarios.getByName(name), raise_err=True)

            scenario.apply()
            return scenario
        except Exception as e:
            raise Exception("Scenario could not be applied:") from e

    # endregion ------------ scenarios ---------------------------------

    # region --------------- data pilot methods ------------------------

    @staticmethod
    def get_pilot_tables(sheet: XSpreadsheet) -> XDataPilotTables:
        """
        Gets pivot tables (formerly known as DataPilot) for a sheet.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet

        Raises:
            MissingInterfaceError: If a required interface is missing.

        Returns:
            XDataPilotTables: Pivot tables
        """
        db_supp = mLo.Lo.qi(XDataPilotTablesSupplier, sheet, raise_err=True)
        return db_supp.getDataPilotTables()

    get_pivot_tables = get_pilot_tables

    @staticmethod
    def get_pilot_table(dp_tables: XDataPilotTables, name: str) -> XDataPilotTable:
        """
        Get a pivot table (formerly known as DataPilot) from a XDataPilotTables instance.

        |lo_safe|

        Args:
            dp_tables (XDataPilotTables): Instance that contains the table.
            name (str): Name of the table to get.

        Raises:
            Exception: If table is not found or other error has occurred.

        Returns:
            XDataPilotTable: Table.
        """
        try:
            table = dp_tables.getByName(name)
            if table is None:
                raise Exception(f"Did not find data pilot table '{name}'")
            return mLo.Lo.qi(XDataPilotTable, table, raise_err=True)
        except Exception as e:
            raise Exception(f"Pilot table lookup failed for '{name}'") from e

    get_pivot_table = get_pilot_table
    # endregion ------------ data pilot methods ------------------------

    # region --------------- using calc functions ----------------------

    @classmethod
    def compute_function(cls, fn: GeneralFunction | str, cell_range: XCellRange) -> float:
        """
        Computes a Calc Function.

        |lo_safe|

        Args:
            fn (GeneralFunction | str): Function to calculate, GeneralFunction Enum value or String such as 'SUM' or 'MAX'.
            cell_range (XCellRange): Cell range to apply function on.

        Returns:
            float: result of function if successful. If there is an error then 0.0 is returned.

        See Also:
            :ref:`ch23_gen_func`
        """
        try:
            sheet_op = mLo.Lo.qi(XSheetOperation, cell_range, raise_err=True)
            func = GeneralFunction(fn)  # convert to enum value if str
            if not isinstance(fn, uno.Enum):
                mLo.Lo.print("Arg fn is invalid, returning 0.0")
                return 0.0
            return sheet_op.computeFunction(func)  # type: ignore
        except Exception as e:
            mLo.Lo.print("Compute function failed. Returning 0.0")
            mLo.Lo.print(f"    {e}")
        return 0.0

    @staticmethod
    def call_fun(func_name: str, *args: Any) -> Any:
        """
        Execute a Calc function by its (English) name and based on the given arguments.

        |lo_unsafe|

        Args:
            func_name (str): the English name of the function to execute
            args: (Any): the arguments of the called function.
                Each argument must be either a string, a numeric value
                or a sequence of sequences ( tuples or list ) combining those types.

        Returns:
            Any: The (string or numeric) value or the array of arrays returned by the call to the function
                When the arguments contain arrays, the function is executed as an array function
                Wrong arguments generate an error
        """
        args_len = len(args)
        arg = () if args_len == 0 else args
        try:
            fa = mLo.Lo.create_instance_mcf(XFunctionAccess, "com.sun.star.sheet.FunctionAccess", raise_err=True)
            return fa.callFunction(func_name.upper(), arg)
        except Exception as e:
            mLo.Lo.print(f"Could not invoke function '{func_name.upper()}'")
            mLo.Lo.print(f"    {e}")
        return None

    @staticmethod
    def get_function_names() -> List[str] | None:
        """
        Get a list of all function names.

        |lo_unsafe|

        Returns:
            List[str] | None: List of function names if found; Otherwise, ``None``
        """
        func_desc = mLo.Lo.create_instance_mcf(XFunctionDescriptions, "com.sun.star.sheet.FunctionDescriptions")
        if func_desc is None:
            mLo.Lo.print("No function descriptions were found")
            return None

        nms: List[str] = []
        for i in range(func_desc.getCount()):
            try:
                props = cast(Sequence[PropertyValue], func_desc.getByIndex(i))
                for p in props:
                    if p.Name == "Name":
                        nms.append(str(p.Value))
                        break
            except Exception:
                continue
        if not nms:
            mLo.Lo.print("No function names were found")
            return None
        nms.sort()
        return nms

    # region    find_function()

    @staticmethod
    def _find_function_by_name(func_nm: str) -> Tuple[PropertyValue, ...] | None:
        """LO UN-Safe method"""
        if not func_nm:
            raise ValueError("Invalid arg, please supply a function name to find.")
        try:
            func_desc = mLo.Lo.create_instance_mcf(
                XFunctionDescriptions, "com.sun.star.sheet.FunctionDescriptions", raise_err=True
            )
        except Exception as e:
            raise Exception("No function descriptions were found") from e

        for i in range(func_desc.getCount()):
            try:
                props = cast(Sequence[PropertyValue], func_desc.getByIndex(i))
                for p in props:
                    if p.Name == "Name" and str(p.Value) == func_nm:
                        return tuple(props)
            except Exception:
                continue
        mLo.Lo.print(f"Function '{func_nm}' not found")
        return None

    @staticmethod
    def _find_function_by_idx(idx: int) -> Tuple[PropertyValue] | None:
        """LO UN-Safe method"""
        # sourcery skip: raise-specific-error
        if idx < 0:
            raise IndexError("Negative index in not allowed.")
        try:
            func_desc = mLo.Lo.create_instance_mcf(
                XFunctionDescriptions, "com.sun.star.sheet.FunctionDescriptions", raise_err=True
            )
        except Exception as e:
            raise Exception("No function descriptions were found") from e

        try:
            return cast(Tuple[PropertyValue], func_desc.getByIndex(idx))
        except Exception as e:
            mLo.Lo.print(f"Could not access function description {idx}")
            mLo.Lo.print(f"    {e}")
        return None

    @overload
    @classmethod
    def find_function(cls, func_nm: str) -> Tuple[PropertyValue] | None:
        """
        Finds a function.

        |lo_unsafe|

        Args:
            func_nm (str): function name.

        Returns:
            Tuple[PropertyValue] | None: Function properties as tuple on success; Otherwise, ``None``.
        """
        ...

    @overload
    @classmethod
    def find_function(cls, idx: int) -> Tuple[PropertyValue] | None:
        """
        Finds a function.

        |lo_unsafe|

        Args:
            idx (int): Index of function.

        Returns:
            Tuple[PropertyValue] | None: Function properties as tuple on success; Otherwise, ``None``.
        """
        ...

    @classmethod
    def find_function(cls, *args, **kwargs) -> Tuple[PropertyValue, ...] | None:
        """
        Finds a function.

        |lo_unsafe|

        Args:
            func_nm (str): function name.
            idx (int): Index of function.

        Returns:
            Tuple[PropertyValue, ...] | None: Function properties as tuple on success; Otherwise, ``None``.
        """
        ordered_keys = (1,)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("func_nm", "idx")
            check = all(key in valid_keys for key in kwargs)
            if not check:
                raise TypeError("find_function() got an unexpected keyword argument")
            keys = ("func_nm", "idx")
            for key in keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            return ka

        if count != 1:
            raise TypeError("find_function() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        if isinstance(kargs[1], int):
            return cls._find_function_by_idx(kargs[1])
        return cls._find_function_by_name(kargs[1])

    # endregion find_function()

    @classmethod
    def print_function_info(cls, func_name: str) -> None:
        """
        Prints Function Info to console.

        |lo_unsafe|

        Args:
            func_name (str): Function name

        Returns:
            None:

        .. versionchanged:: 0.6.10

            Removed cancel event args.
        """
        prop_vals = cls._find_function_by_name(func_nm=func_name)
        if prop_vals is None:
            return
        mProps.Props.show_props(func_name, prop_vals)  # type: ignore
        cls.print_fun_arguments(prop_vals)
        print()

    @classmethod
    def print_fun_arguments(cls, prop_vals: Sequence[PropertyValue]) -> None:
        """
        Prints Function Arguments to console.

        |lo_safe|

        Args:
            prop_vals (Sequence[PropertyValue]): Property values

        Returns:
            None:

        .. versionchanged:: 0.6.10

            Removed cancel event args.
        """
        args = cast("Sequence[FunctionArgument]", mProps.Props.get_value(name="Arguments", props=prop_vals))
        if args is None:
            print("No arguments found")
            return

        print(f"No. of arguments: {len(args)}")
        for i, fa in enumerate(args):
            print(f"{i+1}. Argument name: {fa.Name}")
            print(f"  Description: '{fa.Description}'")
            print(f"  Is optional?: {fa.IsOptional}")
            print()

    @staticmethod
    def get_recent_functions() -> Tuple[int, ...] | None:
        """
        Gets recent functions.

        |lo_unsafe|

        Returns:
            Tuple[int, ...] | None: Tuple of integers that point to functions
        """
        recent_func = mLo.Lo.create_instance_mcf(XRecentFunctions, "com.sun.star.sheet.RecentFunctions")
        if recent_func is None:
            mLo.Lo.print("No recent functions found")
            return None

        return recent_func.getRecentFunctionIds()  # type: ignore

    # endregion ------------ using calc functions ----------------------

    # region --------------- solver methods ----------------------------

    @classmethod
    def goal_seek(
        cls,
        gs: XGoalSeek,
        sheet: XSpreadsheet,
        cell_name: str | mCellObj.CellObj,
        formula_cell_name: str | mCellObj.CellObj,
        result: int | float,
    ) -> float:
        """
        Calculates a value which gives a specified result in a formula.

        |lo_safe|

        Args:
            gs (XGoalSeek): Goal seeking value for cell
            sheet (XSpreadsheet): Spreadsheet
            cell_name (str | CellObj): cell name
            formula_cell_name (str | CellObj): formula cell name
            result (int, float): float or int, result of the goal seek

        Raises:
            GoalDivergenceError: If goal divergence is greater than 0.1

        Returns:
            float: result of the goal seek
        """
        pos = cls._get_cell_address_sheet(sheet=sheet, cell_name=str(cell_name))
        formula_pos = cls._get_cell_address_sheet(sheet=sheet, cell_name=str(formula_cell_name))

        goal_result = gs.seekGoal(formula_pos, pos, f"{float(result)}")
        if goal_result.Divergence >= 0.1:
            mLo.Lo.print(f"NO result; divergence: {goal_result.Divergence}")
            raise mEx.GoalDivergenceError(goal_result.Divergence)
        return goal_result.Result

    @staticmethod
    def get_solvers() -> Tuple[str, ...]:
        """
        Get solvers.

        |lo_unsafe|

        Returns:
            None:

        .. versionadded:: 0.9.1
        """
        nms = mInfo.Info.get_service_names(service_name="com.sun.star.sheet.Solver")
        return tuple(nms) if nms else ()

    @staticmethod
    def list_solvers() -> None:
        """
        Prints solvers.

        |lo_unsafe|
        """
        print("Services offered by the solver:")
        nms = mInfo.Info.get_service_names(service_name="com.sun.star.sheet.Solver")
        if nms is None:
            print("  none")
            return

        for service in nms:
            print(f"  {service}")
        print()

    @staticmethod
    def to_constraint_op(op: str) -> SolverConstraintOperator:
        """
        Convert string value to SolverConstraintOperator.

        If ``op`` is not valid then SolverConstraintOperator.EQUAL is returned.

        |lo_safe|

        Args:
            op (str): Operator such as =, ==, <=, =<, >=, =>

        Returns:
            SolverConstraintOperator: Operator as enum
        """
        if op in {"=", "=="}:
            return SolverConstraintOperator.EQUAL
        if op in {"<=", "=<"}:
            return SolverConstraintOperator.LESS_EQUAL
        if op in {">=", "=>"}:
            return SolverConstraintOperator.GREATER_EQUAL
        mLo.Lo.print(f"Do not recognize op: {op}; using EQUAL")
        return SolverConstraintOperator.EQUAL

    # region    make_constraint()
    @classmethod
    def _make_constraint_op_str_sht_cell_name(
        cls, num: int | float, op: str, sheet: XSpreadsheet, cell_name: str
    ) -> SolverConstraint:
        """LO Safe Method"""
        return cls._make_constraint_op_str_addr(
            num=num, op=op, addr=cls._get_cell_address_sheet(sheet=sheet, cell_name=cell_name)
        )

    @classmethod
    def _make_constraint_op_str_addr(cls, num: int | float, op: str, addr: CellAddress) -> SolverConstraint:
        """LO Safe Method"""
        return cls._make_constraint_op_sco_addr(num=num, op=cls.to_constraint_op(op), addr=addr)

    @classmethod
    def _make_constraint_op_sco_sht_cell_name(
        cls, num: int | float, op: SolverConstraintOperator, sheet: XSpreadsheet, cell_name: str
    ) -> SolverConstraint:
        """LO Safe Method"""
        return cls._make_constraint_op_sco_addr(
            num=num, op=op, addr=cls._get_cell_address_sheet(sheet=sheet, cell_name=cell_name)
        )

    @classmethod
    def _make_constraint_op_sco_addr(
        cls, num: int | float, op: SolverConstraintOperator, addr: CellAddress
    ) -> SolverConstraint:
        """LO Safe Method"""
        sc = SolverConstraint()
        sc.Left = addr
        sc.Operator = op  # type: ignore
        sc.Right = float(num)
        return sc

    @overload
    @classmethod
    def make_constraint(cls, num: int | float, op: str, addr: CellAddress) -> SolverConstraint:
        """
        Makes a constraint for a solver model.

        |lo_safe|

        Args:
            num (Number): Constraint number such as float or int.
            op (str): Operation such as ``<=``.
            addr (CellAddress): Cell Address.

        Returns:
            SolverConstraint: Solver constraint that can be use in a solver model.
        """
        ...

    @overload
    @classmethod
    def make_constraint(cls, num: int | float, op: SolverConstraintOperator, addr: CellAddress) -> SolverConstraint:
        """
        Makes a constraint for a solver model.

        |lo_safe|

        Args:
            num (Number): Constraint number such as float or int.
            op (SolverConstraintOperator): Operation.
            addr (CellAddress): Cell Address.

        Returns:
            SolverConstraint: Solver constraint that can be use in a solver model.
        """
        ...

    @overload
    @classmethod
    def make_constraint(cls, num: int | float, op: str, sheet: XSpreadsheet, cell_name: str) -> SolverConstraint:
        """
        Makes a constraint for a solver model.

        |lo_safe|

        Args:
            num (Number): Constraint number such as float or int.
            op (str): Operation such as ``<=``.
            cell_name (str): Cell name such as ``A1``.

        Returns:
            SolverConstraint: Solver constraint that can be use in a solver model.
        """
        ...

    @overload
    @classmethod
    def make_constraint(
        cls, num: int | float, op: str, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj
    ) -> SolverConstraint:
        """
        Makes a constraint for a solver model.

        |lo_safe|

        Args:
            num (Number): Constraint number such as float or int.
            op (str): Operation such as ``<=``.
            cell_obj (CellObj): Cell Object.

        Returns:
            SolverConstraint: Solver constraint that can be use in a solver model.
        """
        ...

    @overload
    @classmethod
    def make_constraint(
        cls, num: int | float, op: SolverConstraintOperator, sheet: XSpreadsheet, cell_name: str
    ) -> SolverConstraint:
        """
        Makes a constraint for a solver model.

        |lo_safe|

        Args:
            num (Number): Constraint number such as float or int.
            op (SolverConstraintOperator): Operation.
            sheet (XSpreadsheet): Spreadsheet.
            cell_name (str): Cell name such as ``A1``.

        Returns:
            SolverConstraint: Solver constraint that can be use in a solver model.
        """
        ...

    @classmethod
    def make_constraint(cls, *args, **kwargs) -> SolverConstraint:
        """
        Makes a constraint for a solver model.

        |lo_safe|

        Args:
            num (Number): Constraint number such as float or int.
            op (str | SolverConstraintOperator): Operation such as ``<=``.
            addr (CellAddress): Cell Address.
            cell_name (str): Cell name such as ``A1``.
            cell_obj (CellObj): Cell Object.
            sheet (XSpreadsheet): Spreadsheet.

        Returns:
            SolverConstraint: Solver constraint that can be use in a solver model.
        """
        ordered_keys = (1, 2, 3, 4)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("num", "op", "sheet", "addr", "cell_name", "cell_obj")
            check = all(key in valid_keys for key in kwargs)
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
            keys = ("cell_name", "cell_obj")
            for key in keys:
                if key in kwargs:
                    ka[4] = kwargs[key]
                    break
            return ka

        if count not in (3, 4):
            raise TypeError("make_constraint() got an invalid number of arguments")

        kargs = get_kwargs()

        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        num = cast(float, kargs[1])
        arg2 = kargs[2]
        arg3 = kargs[3]
        if count == 3:
            return (
                cls._make_constraint_op_str_addr(num=num, op=arg2, addr=arg3)
                if isinstance(arg2, str)
                else cls._make_constraint_op_sco_addr(num=num, op=arg2, addr=arg3)
            )
        if isinstance(arg2, str):
            # def make_constraint(num: float, op: str, sheet: XSpreadsheet, cell_name:str)
            return cls._make_constraint_op_str_sht_cell_name(num=num, op=arg2, sheet=arg3, cell_name=str(kargs[4]))
        else:
            # def make_constraint(num: float, op: SolverConstraintOperator, sheet: XSpreadsheet, cell_name:str)
            return cls._make_constraint_op_sco_sht_cell_name(num=num, op=arg2, sheet=arg3, cell_name=str(kargs[4]))

    # endregion    make_constraint()

    @classmethod
    def solver_report(cls, solver: XSolver) -> None:
        """
        Prints the result of solver.

        |lo_safe|

        Args:
            solver (XSolver): Solver to print result of.
        """
        # note: in original java it was getSuccess(), getObjective(), getVariables(), getSolution(),
        # These are typedef properties. The types-unopy typings are correct. Typedef are represented as Class Properties.
        # is_successful = solver.Success
        cell_name = cls._get_cell_str_addr(solver.Objective)
        print("Solver result: ")
        print(f"  {cell_name} == {solver.ResultValue:.4f}")
        addrs = solver.Variables
        solver_solution = solver.Solution
        print("Solver variables: ")
        for i, num in enumerate(solver_solution):
            cell_name = cls._get_cell_str_addr(addrs[i])
            print(f"  {cell_name} == {num:.4f}")
        print()

    # endregion ------------ solver methods ----------------------------

    # region --------------- headers /footers --------------------------

    @staticmethod
    def get_head_foot(props: XPropertySet, content: str) -> XHeaderFooterContent:
        """
        Gets header footer content.

        |lo_safe|

        Args:
            props (XPropertySet): Properties.
            content (str): content.

        Raises:
            MissingInterfaceError: If unable to obtain ``XHeaderFooterContent`` interface.

        Returns:
            XHeaderFooterContent: Header Footer Content

        See Also:
            `LibreOffice API XHeaderFooterContent <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1sheet_1_1XHeaderFooterContent.html>`_
        """
        return mLo.Lo.qi(XHeaderFooterContent, mProps.Props.get(props, content), raise_err=True)

    @staticmethod
    def print_head_foot(title: str, hfc: XHeaderFooterContent) -> None:
        """
        Prints header, footer to console.

        |lo_safe|

        Args:
            title (str): Title printed to console.
            hfc (XHeaderFooterContent): Content.

        Returns:
            None:

        .. versionchanged:: 0.6.10

            Removed cancel event args.
        """
        left = hfc.getLeftText()
        center = hfc.getCenterText()
        right = hfc.getRightText()
        print(f"{title}: '{left.getString()}' : '{center.getString()}' : '{right.getString()}'")

    @classmethod
    def get_region(cls, hfc: XHeaderFooterContent, region: Calc.HeaderFooter) -> XText:
        """
        Get region for Header-Footer-Content.

        |lo_safe|

        Args:
            hfc (XHeaderFooterContent): Content.
            region (HeaderFooter): Region to get.

        Raises:
            TypeError: If hfc or region is not a valid type.

        Returns:
            XText: interface instance
        """
        if hfc is None:
            raise TypeError("'hfc' is expected to be XHeaderFooterContent instance")

        if region == cls.HeaderFooter.HF_LEFT:
            return hfc.getLeftText()
        if region == cls.HeaderFooter.HF_CENTER:
            return hfc.getCenterText()
        if region == cls.HeaderFooter.HF_RIGHT:
            return hfc.getRightText()
        raise TypeError("region is not a valid type")

    @classmethod
    def set_head_foot(cls, hfc: XHeaderFooterContent, region: Calc.HeaderFooter, text: str) -> None:
        """
        Sets the Header-Footer-Content.

        |lo_safe|

        Args:
            hfc (XHeaderFooterContent): Content.
            region (HeaderFooter): Region to set.
            text (str): Text to set.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_page_header`
            - :ref:`help_calc_format_modify_page_footer`
        """
        xtext = cls.get_region(hfc=hfc, region=region)
        if xtext is None:
            mLo.Lo.print("Could not set text")
            return
        header_cursor = xtext.createTextCursor()
        header_cursor.gotoStart(False)
        header_cursor.gotoEnd(True)
        header_cursor.setString(text)

    # endregion --------------- headers /footers -----------------------

    # region --------------- styles ------------------------------------
    # region set_style_range()
    @overload
    @classmethod
    def set_style_range(cls, sheet: XSpreadsheet, range_name: str, *, styles: Sequence[StyleT]) -> None:
        """
        Set style/formatting on cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_name (str): Range Name such as ``A1:D5``.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_style_range(cls, sheet: XSpreadsheet, range_obj: mRngObj.RangeObj, *, styles: Sequence[StyleT]) -> None:
        """
        Set style/formatting on cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_obj (RangeObj): Range Object.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_style_range(cls, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj, *, styles: Sequence[StyleT]) -> None:
        """
        Set style/formatting on cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            cell_obj (CellObj): Cell Object.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_style_range(cls, sheet: XSpreadsheet, cr_addr: CellRangeAddress, *, styles: Sequence[StyleT]) -> None:
        """
        Set style/formatting on cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            cr_addr (CellRangeAddress): Cell range Address.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_style_range(cls, cell_range: XCellRange, *, styles: Sequence[StyleT]) -> None:
        """
        Set style/formatting on cell range.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range. If passed in then the same instance is returned.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_style_range(
        cls,
        sheet: XSpreadsheet,
        col_start: int,
        row_start: int,
        col_end: int,
        row_end: int,
        *,
        styles: Sequence[StyleT],
    ) -> None:
        """
        Set style/formatting on cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @classmethod
    def set_style_range(cls, *args, **kwargs) -> None:
        """
        Set style/formatting on cell range.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet Document.
            range_name (str): Range Name such as ``A1:D5``.
            range_obj (RangeObj): Range Object.
            cell_obj (CellObj): Cell Object.
            cr_addr (CellRangeAddress): Cell range Address.
            cell_range (XCellRange): Cell Range. If passed in then the same instance is returned.
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_style_cell`
            - :ref:`help_calc_format_direct_cell`

        .. versionadded:: 0.9.2
        """
        styles = cast(Sequence[StyleT], kwargs.pop("styles", None))
        if styles is None:
            return
        cell_range = cls.get_cell_range(*args, **kwargs)
        supported_styles = (
            "com.sun.star.style.CharacterProperties",
            "com.sun.star.style.ParagraphProperties",
            "com.sun.star.table.CellProperties",
            "com.sun.star.sheet.SheetCellRange",
            "com.sun.star.sheet.SheetCell",
        )

        for style in styles:
            if style.support_service(*supported_styles):
                style.apply(cell_range)

    # endregion set_style_range()

    # region set_style_cell()
    @overload
    @classmethod
    def set_style_cell(cls, sheet: XSpreadsheet, addr: CellAddress, *, styles: Sequence[StyleT]) -> None:
        """
        Set style/formatting on cell.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            addr (CellAddress): Cell Address.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_style_cell(cls, sheet: XSpreadsheet, cell_name: str, *, styles: Sequence[StyleT]) -> None:
        """
        Set style/formatting on cell.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_name (str): Cell Name such as ``A1``.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_style_cell(cls, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj, *, styles: Sequence[StyleT]) -> None:
        """
        Set style/formatting on cell.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            cell_obj: (CellObj): Cell object.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_style_cell(cls, sheet: XSpreadsheet, col: int, row: int, *, styles: Sequence[StyleT]) -> None:
        """
        Set style/formatting on cell.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            col (int): Cell column.
            row (int): cell row.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_style_cell(cls, cell_range: XCellRange, *, styles: Sequence[StyleT]) -> None:
        """
        Set style/formatting on cell.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def set_style_cell(cls, cell_range: XCellRange, col: int, row: int, *, styles: Sequence[StyleT]) -> None:
        """
        Set style/formatting on cell.

        |lo_safe|

        Args:
            cell_range (XCellRange): Cell Range.
            col (int): Cell column.
            row (int): cell row.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @classmethod
    def set_style_cell(cls, *args, **kwargs) -> None:
        """
        Set style/formatting on cell.

        |lo_safe|

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            addr (CellAddress): Cell Address.
            cell_name (str): Cell Name such as ``A1``.
            cell_obj: (CellObj): Cell object.
            cell_range (XCellRange): Cell Range.
            col (int): Cell column.
            row (int): cell row.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_style_cell`
            - :ref:`help_calc_format_direct_cell`

        .. versionadded:: 0.9.2
        """
        styles = cast(Sequence[StyleT], kwargs.pop("styles", None))
        if styles is None:
            return
        cell = cls.get_cell(*args, **kwargs)
        supported_styles = (
            "com.sun.star.style.CharacterProperties",
            "com.sun.star.style.ParagraphProperties",
            "com.sun.star.table.CellProperties",
            "com.sun.star.sheet.SheetCellRange",
            "com.sun.star.sheet.SheetCell",
        )

        for style in styles:
            if style.support_service(*supported_styles):
                style.apply(cell)

    # endregion set_style_cell()
    # endregion ------------ styles ------------------------------------

    # region --------------- dispatch ----------------------------------
    @staticmethod
    def dispatch_recalculate() -> None:
        """
        Dispatches recalculate command to the current sheet.

        |lo_unsafe|

        Also useful when needing to refresh a chart.

        Returns:
            None:

        .. versionadded:: 0.9.4
        """
        mLo.Lo.dispatch_cmd("Calculate")

    # endregion ------------ dispatch ----------------------------------

    # region ------------------- printer methods ---------------------------

    # region print_sheet()
    @overload
    @classmethod
    def print_sheet(
        cls, *, printer_name: str, range_name: str, idx: Optional[int], doc: Optional[XSpreadsheetDocument]
    ) -> None:
        """
        Print a sheet to the specified printer directly.

        |lo_unsafe|

        Args:
            printer_name (str): Name of Printer to use such as "Brother MFC-L2750DW series"
            range_name (str): Range Name such as ``A1:D5``
            idx (int, optional): Index of sheet to print. If not specified then the active sheet is used.
            doc (XSpreadsheetDocument, optional): Document to use. If not specified then the active document is used.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def print_sheet(
        cls, *, printer_name: str, range_obj: mRngObj.RangeObj, idx: Optional[int], doc: Optional[XSpreadsheetDocument]
    ) -> None:
        """
        Print a sheet to the specified printer directly.

        |lo_unsafe|

        Args:
            printer_name (str): Name of Printer to use such as "Brother MFC-L2750DW series"
            range_obj (RangeObj): Range Object
            idx (int, optional): Index of sheet to print. If not specified then the active sheet is used.
            doc (XSpreadsheetDocument, optional): Document to use. If not specified then the active document is used.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def print_sheet(
        cls, *, printer_name: str, cell_obj: mCellObj.CellObj, idx: Optional[int], doc: Optional[XSpreadsheetDocument]
    ) -> None:
        """
        Print a sheet to the specified printer directly.

        |lo_unsafe|

        Args:
            printer_name (str): Name of Printer to use such as "Brother MFC-L2750DW series"
            cell_obj (CellObj): Cell Object
            idx (int, optional): Index of sheet to print. If not specified then the active sheet is used.
            doc (XSpreadsheetDocument, optional): Document to use. If not specified then the active document is used.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def print_sheet(
        cls, *, printer_name: str, cr_addr: CellRangeAddress, idx: Optional[int], doc: Optional[XSpreadsheetDocument]
    ) -> None:
        """
        Print a sheet to the specified printer directly.

        |lo_unsafe|

        Args:
            printer_name (str): Name of Printer to use such as "Brother MFC-L2750DW series"
            cr_addr (CellRangeAddress): Cell range Address
            idx (int, optional): Index of sheet to print. If not specified then the active sheet is used.
            doc (XSpreadsheetDocument, optional): Document to use. If not specified then the active document is used.

        Returns:
            None:
        """
        ...

    @overload
    @classmethod
    def print_sheet(cls, *, printer_name: str, cell_range: XCellRange, doc: Optional[XSpreadsheetDocument]) -> None:
        """
        Print a sheet to the specified printer directly.

        |lo_unsafe|

        Args:
            printer_name (str): Name of Printer to use such as "Brother MFC-L2750DW series"
            cell_range (XCellRange): Cell Range. If passed in then the same instance is returned.
            doc (XSpreadsheetDocument, optional): Document to use. If not specified then the active document is used.

        Returns:
            None:
        """
        ...

    @classmethod
    def print_sheet(cls, *, printer_name: str, **kwargs) -> None:
        """
        Print a sheet to the specified printer directly.

        |lo_unsafe|

        Args:
            printer_name (str): Name of Printer to use such as "Brother MFC-L2750DW series"
            idx (int, optional): Index of sheet to print. If not specified then the active sheet is used.
            range_name (str): Range Name such as ``A1:D5``
            range_obj (RangeObj): Range Object
            cell_obj (CellObj): Cell Object
            cr_addr (CellRangeAddress): Cell range Address
            cell_range (XCellRange): Cell Range. If passed in then the same instance is returned.
            doc (XSpreadsheetDocument, optional): Document to use. If not specified then the active document is used.

        Raises:
            ValueError: If printer_name is not specified
            ValueError: If no range is specified

        Returns:
            None:

        See Also:
            - :py:meth:`~.calc.Calc.get_selected_addr`
            - :ref:`help_calc_module_class_print_sheet`

        .. versionadded:: 0.11.2
        """
        # sourcery skip: assign-if-exp, merge-else-if-into-elif
        if not printer_name:
            raise ValueError("No printer name specified")
        idx = int(kwargs.get("idx", -1))
        doc = kwargs.get("doc", None)
        print_props = mProps.Props.make_props(Name=printer_name)

        if doc is None:
            if idx < 0:
                sheet = cls.get_active_sheet()
            else:
                sheet = cls.get_sheet(idx=idx)
        else:
            if idx < 0:
                sheet = cls.get_active_sheet(doc=doc)
            else:
                sheet = cls.get_sheet(doc=doc, idx=idx)

        if "range_name" in kwargs:
            cell_range = cls.get_cell_range(sheet=sheet, range_name=kwargs["range_name"])
        elif "range_obj" in kwargs:
            cell_range = cls.get_cell_range(sheet=sheet, range_obj=kwargs["range_obj"])
        elif "cell_obj" in kwargs:
            cell_range = cls.get_cell_range(sheet=sheet, cell_obj=kwargs["cell_obj"])
        elif "cr_addr" in kwargs:
            cell_range = cls.get_cell_range(sheet=sheet, cr_addr=kwargs["cr_addr"])
        elif "cell_range" in kwargs:
            cell_range = cls.get_cell_range(cell_range=kwargs["cell_range"])
        else:
            raise ValueError("No range specified")

        addressable = mLo.Lo.qi(XCellRangeAddressable, cell_range, True)

        # remove all configured print areas
        if doc is None:
            sheets = cls.get_sheets()
        else:
            sheets = cls.get_sheets(doc=doc)
        enum_access = mLo.Lo.qi(XEnumerationAccess, sheets, True)
        enums = enum_access.createEnumeration()
        while enums.hasMoreElements():
            sh = mLo.Lo.qi(XPrintAreas, enums.nextElement())
            if sh is not None:
                sh.setPrintAreas(())

        # set single print area
        sh = mLo.Lo.qi(XPrintAreas, sheet, True)
        sh.setPrintAreas((addressable.getRangeAddress(),))
        printable = mLo.Lo.qi(XPrintable, doc, True)
        printable.print(print_props)

    # endregion print_sheet()


# endregion --------------- printer methods ----------------------------
