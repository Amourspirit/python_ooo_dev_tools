from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING

from ooodev.format.proto.calc.numbers.numbers_t import NumbersT as CalcNumbersT

from ooodev.mock.mock_g import DOCS_BUILDING

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Self

    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
    from com.sun.star.chart2 import XChartDocument
    from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum
    from ooo.dyn.lang.locale import Locale
    from ooo.dyn.util.number_format import NumberFormatEnum
else:
    Protocol = object
    Self = Any
    XChartDocument = Any
    ChartNumbersT = Any
    NumberFormatIndexEnum = Any
    Locale = Any
    NumberFormatEnum = Any


class NumbersT(CalcNumbersT, Protocol):
    """Numbers Protocol"""

    def __init__(
        self,
        chart_doc: XChartDocument,
        *,
        num_format: NumberFormatEnum | int = ...,
        num_format_index: NumberFormatIndexEnum | int = ...,
        lang_locale: Locale | None = ...,
    ) -> None:
        """
        Constructor

        Args:
            chart_doc (XChartDocument): Chart document.
            num_format (NumberFormatEnum, int, optional): Type of a number format.
                Use this to select a default format.
                Only used if ``num_format_index`` is ``-1`` (omitted).
            num_format_index (NumberFormatIndexEnum | int, optional): Index of a number format.
                The enumeration values represent the built-in number formats.
            lang_locale (Locale, optional): Locale of the number format.
        """
        ...

    @classmethod
    def from_str(
        cls,
        chart_doc: XChartDocument,
        nf_str: str,
        lang_locale: Locale | None = None,
        auto_add: bool = False,
        **kwargs,
    ) -> NumbersT: ...

    @classmethod
    def from_index(
        cls, chart_doc: XChartDocument, index: int, lang_locale: Locale | None = None, **kwargs
    ) -> NumbersT: ...

    @overload
    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: Any) -> NumbersT: ...

    @overload
    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: Any, **kwargs) -> NumbersT: ...

    # region Instance Properties

    @property
    def general(self) -> Self:
        """Gets general format"""
        ...

    @property
    def date(self) -> Self:
        """Gets date format"""
        ...

    @property
    def time(self) -> Self:
        """Gets time format"""
        ...

    @property
    def currency(self) -> Self:
        """Gets currency format"""
        ...

    @property
    def number(self) -> Self:
        """Gets number format"""
        ...

    @property
    def scientific(self) -> Self:
        """Gets scientific format"""
        ...

    @property
    def fraction(self) -> Self:
        """Gets fraction format"""
        ...

    @property
    def percent(self) -> Self:
        """Gets percent format"""
        ...

    @property
    def text(self) -> Self:
        """Gets text format"""
        ...

    @property
    def datetime(self) -> Self:
        """Gets datetime format"""
        ...

    @property
    def boolean(self) -> Self:
        """Gets boolean format"""
        ...

    # endregion Instance Properties
