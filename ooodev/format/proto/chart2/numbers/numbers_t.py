from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING

from ooodev.format.proto.calc.numbers.numbers_t import NumbersT as CalcNumbersT

from ooodev.mock.mock_g import DOCS_BUILDING

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Self

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
    ) -> NumbersT:
        """
        Gets instance from format string

        Args:
            chart_doc (XChartDocument): Chart document.
            nf_str (str): Format string.
            lang_locale (Locale, optional): Locale. Defaults to ``None``.
            auto_add (bool, optional): If True, format string will be added to document if not found. Defaults to ``False``.

        Keyword Args:
            source_format (bool, optional): If ``True``, the number format will be linked to the source format. Defaults to ``False``.

        Returns:
            NumbersT: Instance that represents numbers format.
        """
        ...

    @classmethod
    def from_index(
        cls, chart_doc: XChartDocument, index: int, lang_locale: Locale | None = None, **kwargs
    ) -> NumbersT:
        """
        Gets instance from number format index. This is the index that is assigned to the ``NumberFormat`` property of an object such as a cell.

        Args:
            chart_doc (XChartDocument): Chart document.
            index (int): Format (``NumberFormat``) index.
            lang_locale (Locale, optional): Locale. Defaults to ``None``.

        Keyword Args:
            source_format (bool, optional): If ``True``, the number format will be linked to the source format. Defaults to ``False``.

        Returns:
            NumbersT: Instance that represents numbers format.
        """
        ...

    @overload
    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: Any) -> NumbersT:
        """
        Gets instance from object

        Args:
            chart_doc (XChartDocument): Chart document.
            obj (object): UNO Object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            NumberFormat: Instance that represents numbers format.
        """
        ...

    @overload
    @classmethod
    def from_obj(cls, chart_doc: XChartDocument, obj: Any, **kwargs) -> NumbersT:
        """
        Gets instance from object

        Args:
            chart_doc (XChartDocument): Chart document.
            obj (object): UNO Object.
            kwargs (dict): Keyword arguments.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            NumberFormat: Instance that represents numbers format.
        """
        ...

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
