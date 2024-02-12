from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING


if TYPE_CHECKING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
    from ooodev.format.proto.style_t import StyleT
    from com.sun.star.lang import XComponent
    from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum
    from ooo.dyn.lang.locale import Locale
    from ooo.dyn.util.number_format import NumberFormatEnum
else:
    Protocol = object
    StyleT = Any
    XComponent = Any
    NumberFormatIndexEnum = Any
    Locale = Any
    NumberFormatEnum = Any


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
    def from_str(
        cls, nf_str: str, lang_locale: Locale | None = None, auto_add: bool = False, **kwargs
    ) -> NumbersT: ...

    @classmethod
    def from_index(cls, index: int, lang_locale: Locale | None = None, **kwargs) -> NumbersT: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> NumbersT: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> NumbersT: ...

    # region Instance Properties

    @property
    def general(self) -> NumbersT:
        """Gets general format"""
        ...

    @property
    def date(self) -> NumbersT:
        """Gets date format"""
        ...

    @property
    def time(self) -> NumbersT:
        """Gets time format"""
        ...

    @property
    def currency(self) -> NumbersT:
        """Gets currency format"""
        ...

    @property
    def number(self) -> NumbersT:
        """Gets number format"""
        ...

    @property
    def scientific(self) -> NumbersT:
        """Gets scientific format"""
        ...

    @property
    def fraction(self) -> NumbersT:
        """Gets fraction format"""
        ...

    @property
    def percent(self) -> NumbersT:
        """Gets percent format"""
        ...

    @property
    def text(self) -> NumbersT:
        """Gets text format"""
        ...

    @property
    def datetime(self) -> NumbersT:
        """Gets datetime format"""
        ...

    @property
    def boolean(self) -> NumbersT:
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
