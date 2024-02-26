from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING


from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Self
    from typing_extensions import Protocol
    from com.sun.star.lang import XComponent
    from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum
    from ooo.dyn.lang.locale import Locale
    from ooo.dyn.util.number_format import NumberFormatEnum
else:
    Protocol = object
    Self = Any
    XComponent = Any
    NumberFormatIndexEnum = Any
    Locale = Any
    NumberFormatEnum = Any


# see ooodev.format.inner.direct.calc.numbers.numbers.Numbers
class NumbersT(StyleT, Protocol):
    """Numbers Protocol"""

    def __init__(
        self,
        num_format: NumberFormatEnum | int = ...,
        num_format_index: NumberFormatIndexEnum | int = ...,
        lang_locale: Locale | None = ...,
        component: XComponent | None = ...,
    ) -> None:
        """
        Constructor

        Args:
            num_format (NumberFormatEnum, int, optional): Type of a number format.
                Use this to select a default format. Defaults to 0 (General Format).
                Only used if ``num_format_index`` is ``-1`` (omitted).
            num_format_index (NumberFormatIndexEnum, int, optional): Index of a number format.
                The enumeration values represent the built-in number formats. Defaults to ``-1``.
            lang_locale (Locale, optional): Locale of the number format. Defaults to ``None`` which used current Locale.
            component (XComponent, optional): Document such as Spreadsheet or Chart. If Omitted, the current document is used. Defaults to ``None``.

        Returns:
            None:
        """
        ...

    @classmethod
    def from_str(cls, nf_str: str, lang_locale: Locale | None = None, auto_add: bool = False, **kwargs) -> NumbersT:
        """
        Gets instance from format string

        Args:
            nf_str (str): Format string.
            lang_locale (Locale, optional): Locale. Defaults to ``None``.
            auto_add (bool, optional): If True, format string will be added to document if not found. Defaults to ``False``.

        Keyword Args:
            component (XComponent): Calc document. Default is current document.

        Returns:
            Numbers: Instance that represents numbers format.
        """
        ...

    @classmethod
    def from_index(cls, index: int, lang_locale: Locale | None = None, **kwargs) -> NumbersT:
        """
        Gets instance from number format index. This is the index that is assigned to the ``NumberFormat`` property of an object such as a cell.

        Args:
            index (int): Format (``NumberFormat``) index.
            lang_locale (Locale, optional): Locale. Defaults to ``None``.

        Keyword Args:
            component (XComponent): Calc document. Default is current document.

        Returns:
            Numbers: Instance that represents numbers format.
        """
        ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> NumbersT: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> NumbersT:
        """
        Gets instance from object

        Args:
            obj (object): UNO Object.

        Keyword Args:
            component (XComponent): Calc document. Default is current document.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Numbers: Instance that represents numbers format.
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

    # region Properties

    @property
    def prop_format_key(self) -> int:
        """Gets the format key"""
        ...

    @property
    def prop_format_str(self) -> str:
        """Gets the format string, e.g. ``#,##0.00``"""
        ...

    # endregion Properties
