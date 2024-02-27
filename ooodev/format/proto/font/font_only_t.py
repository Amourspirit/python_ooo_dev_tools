from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING
import uno

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_multi_t import StyleMultiT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Self
    from typing_extensions import Protocol
    from ooodev.format.proto.font.font_lang_t import FontLangT
    from ooodev.units.unit_obj import UnitT
    from ooodev.units.unit_pt import UnitPT
else:
    Protocol = object
    Self = Any
    FontLangT = Any
    UnitT = Any
    UnitPT = Any


class FontOnlyT(StyleMultiT, Protocol):
    """Font Only Protocol"""

    def __init__(
        self,
        *,
        name: str | None = ...,
        size: float | UnitT | None = ...,
        font_style: str | None = ...,
        lang: FontLangT | None = ...,
    ) -> None: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> FontOnlyT: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> FontOnlyT: ...

    # region Format Methods
    def fmt_size(self, value: float | UnitT | None = None) -> Self:
        """
        Get copy of instance with text size set.

        Args:
            value (float, UnitT, optional): The size of the characters in ``pt`` (point) units or :ref:`proto_unit_obj`.

        Returns:
            FontOnly: Font with style added
        """
        ...

    def fmt_name(self, value: str | None = None) -> Self:
        """
        Get copy of instance with name set.

        Args:
            value (str, optional): The name of the font.

        Returns:
            FontOnly: Font with style added or removed
        """
        ...

    def fmt_style_name(self, value: str | None = None) -> Self:
        """
        Get copy of instance with style name set.

        Args:
            value (str, optional): The style name of the font.

        Returns:
            FontOnly: Font with style added or removed
        """
        ...

    # endregion Format Methods

    # region Style Properties
    @property
    def prop_size(self) -> UnitPT | None:
        """This value contains the size of the characters in point."""
        ...

    @prop_size.setter
    def prop_size(self, value: float | UnitT | None) -> None: ...

    @property
    def prop_name(self) -> str | None:
        """This property specifies the name of the font style."""
        ...

    @prop_name.setter
    def prop_name(self, value: str | None) -> None: ...

    @property
    def prop_style_name(self) -> str | None:
        """This property specifies the style name of the font style."""
        ...

    @prop_style_name.setter
    def prop_style_name(self, value: str | None) -> None:
        # style name will be added or removed in _set_fd_style()
        ...

    @property
    def prop_inner(self) -> FontLangT:
        """Gets Lang instance"""
        ...

    # endregion Prop Properties
