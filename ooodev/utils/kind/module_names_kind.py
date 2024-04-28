from __future__ import annotations
from typing import Any
from enum import Enum


class ModuleNamesKind(Enum):
    """
    Enum for looking up module names.

    Support conversion from int, str, and tuple.

    Example:
        .. code-block:: python

            >>> from ooodev.utils.kind.module_names_kind import ModuleNamesKind
            # ...
            >>> val1 = ModuleNamesKind.SPREADSHEET_DOCUMENT
            >>> val2 = ModuleNamesKind(20)
            >>> val3 = ModuleNamesKind("com.sun.star.sheet.SpreadsheetDocument")
            >>> val4 = ModuleNamesKind((20, "com.sun.star.sheet.SpreadsheetDocument"))
            >>> val1 == val2 == val3 == val4
            True
    """

    NONE = (0, "unknown")
    FORM_DESIGN = (1, "com.sun.star.sdb.FormDesign")
    VIEW_DESIGN = (2, "com.sun.star.sdb.ViewDesign")
    BASIC_IDE = (3, "com.sun.star.script.BasicIDE")
    QUERY_DESIGN = (4, "com.sun.star.sdb.QueryDesign")
    TABLE_DESIGN = (5, "com.sun.star.sdb.TableDesign")
    WEB_DOCUMENT = (6, "com.sun.star.text.WebDocument")
    START_MODULE = (7, "com.sun.star.frame.StartModule")
    TABLE_DATA_VIEW = (8, "com.sun.star.sdb.TableDataView")
    TEXT_DOCUMENT = (9, "com.sun.star.text.TextDocument")
    BIBLIOGRAPHY = (10, "com.sun.star.frame.Bibliography")
    RELATION_DESIGN = (11, "com.sun.star.sdb.RelationDesign")
    GLOBAL_DOCUMENT = (12, "com.sun.star.text.GlobalDocument")
    CHART_DOCUMENT = (13, "com.sun.star.chart2.ChartDocument")
    TEXT_REPORT_DESIGN = (14, "com.sun.star.sdb.TextReportDesign")
    DATA_SOURCE_BROWSER = (15, "com.sun.star.sdb.DataSourceBrowser")
    XML_FORM_DOCUMENT = (16, "com.sun.star.xforms.XMLFormDocument")
    DRAWING_DOCUMENT = (17, "com.sun.star.drawing.DrawingDocument")
    REPORT_DEFINITION = (18, "com.sun.star.report.ReportDefinition")
    FORMULA_PROPERTIES = (19, "com.sun.star.formula.FormulaProperties")
    SPREADSHEET_DOCUMENT = (20, "com.sun.star.sheet.SpreadsheetDocument")
    OFFICE_DATABASE_DOCUMENT = (21, "com.sun.star.sdb.OfficeDatabaseDocument")
    PRESENTATION_DOCUMENT = (22, "com.sun.star.presentation.PresentationDocument")

    def __str__(self) -> str:
        """Gets the string portion of the enum tuple value."""
        return self.value[1]

    def __int__(self) -> int:
        """Gets the int portion of the enum tuple value."""
        return self.value[0]

    def to_json(self) -> int:
        """Gets the JSON representation of the enum."""
        return self.value[0]


def _new_module_names_kind(cls, val: Any) -> ModuleNamesKind:
    if isinstance(val, ModuleNamesKind):
        return val
    if isinstance(val, str):
        for item in ModuleNamesKind:
            if val == item.value[1]:
                return item
    elif isinstance(val, int):
        for item in ModuleNamesKind:
            if val == item.value[0]:
                return item
    elif isinstance(val, tuple):
        for item in ModuleNamesKind:
            if val == item.value:
                return item
    raise ValueError(f"Invalid value for ModuleNamesKind: {val}")


setattr(ModuleNamesKind, "__new__", _new_module_names_kind)
