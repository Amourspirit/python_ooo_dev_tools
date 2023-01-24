from __future__ import annotations
from typing import Tuple

from ....events.args.key_val_cancel_args import KeyValCancelArgs
from ....meta.static_prop import static_prop
from ...kind.format_kind import FormatKind
from ...style_base import StyleBase
from .kind.style_list_kind import StyleListKind as StyleListKind


class StyleList(StyleBase):
    """
    Style List. Manages List styles for Writer.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together Border Table properties.

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    def __init__(self, name: StyleListKind | str = "") -> None:
        super().__init__(**{self._get_property_name(): str(name)})

    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.style.ParagraphProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.ParagraphProperties",)

    def _get_property_name(self) -> str:
        return "NumberingStyleName"

    # region Style Properties
    @property
    def none(self) -> StyleList:
        """Gets Style that is set to no list style"""
        return StyleList(StyleListKind.NONE)

    @property
    def list_01(self) -> StyleList:
        """Gets List style 01 (Bullet •)"""
        return StyleList(StyleListKind.LIST_01)

    @property
    def list_02(self) -> StyleList:
        """Gets List style 01 (Dash –)"""
        return StyleList(StyleListKind.LIST_02)

    @property
    def list_03(self) -> StyleList:
        """Gets List style 03 ( 🗹 [checkbox like])"""
        return StyleList(StyleListKind.LIST_03)

    @property
    def list_04(self) -> StyleList:
        """Gets List style 0 (triangle like)"""
        return StyleList(StyleListKind.LIST_04)

    @property
    def list_05(self) -> StyleList:
        """Gets List style 05 (Bullet ꭗ)"""
        return StyleList(StyleListKind.LIST_05)

    @property
    def num_123(self) -> StyleList:
        """Gets List style Numbering 123"""
        return StyleList(StyleListKind.NUM_123)

    @property
    def num_abc(self) -> StyleList:
        """Gets List style Numbering abc (lower case)"""
        return StyleList(StyleListKind.NUM_abc)

    @property
    def num_ABC(self) -> StyleList:
        """Gets List style Numbering ABC (upper case)"""
        return StyleList(StyleListKind.NUM_ABC)

    @property
    def num_ivx(self) -> StyleList:
        """Gets List style Numbering ivx (lower case)"""
        return StyleList(StyleListKind.NUM_ivx)

    @property
    def num_IVX(self) -> StyleList:
        """Gets List style Numbering IVX (upper case)"""
        return StyleList(StyleListKind.NUM_IVX)

    # endregion Style Properties

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.STYLE | FormatKind.STATIC

    @property
    def prop_name(self) -> str:
        """Gets/Sets Character style namd"""
        return self._get(self._get_property_name())

    @prop_name.setter
    def prop_name(self, value: StyleListKind | str) -> None:
        if self is StyleList.default:
            raise ValueError("Setting StyleList.default properties is not allowed.")
        self._set(self._get_property_name(), str(value))

    @static_prop
    def default() -> StyleList:  # type: ignore[misc]
        """Gets ``StyleList`` default. Static Property."""
        if StyleList._DEFAULT is None:
            StyleList._DEFAULT = StyleList(name="")
        return StyleList._DEFAULT
