# region Import
from __future__ import annotations
from typing import Tuple

from ooodev.meta.static_prop import static_prop
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleName
from ooodev.format.writer.style.lst.style_list_kind import StyleListKind as StyleListKind

# endregion Import


class BulletList(StyleName):
    """
    Style List. Manages List styles for Writer.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together Border Table properties.

    .. seealso::

        - :ref:`help_writer_format_style_bullet_list`

    .. versionadded:: 0.9.0
    """

    def __init__(self, name: StyleListKind | str = "") -> None:
        """
        Constructor

        Args:
            name (Any): Style Name.

        Raises:
            ValueError: If Name is ``None`` or empty string.

        Returns:
            None:

        See Also:

            - :ref:`help_writer_format_style_bullet_list`
        """
        super().__init__(name=name)

    # region Overrides
    def _get_family_style_name(self) -> str:
        return "NumberingStyles"

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.ParagraphProperties",)
        return self._supported_services_values

    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "NumberingStyleName"
        return self._property_name

    # endregion Overrides

    # region Style Properties
    @property
    def none(self) -> BulletList:
        """Gets Style that is set to no list style"""
        return BulletList(StyleListKind.NONE)

    @property
    def list_01(self) -> BulletList:
        """Gets List style 01 (Bullet •)"""
        return BulletList(StyleListKind.LIST_01)

    @property
    def list_02(self) -> BulletList:
        """Gets List style 01 (Dash –)"""
        return BulletList(StyleListKind.LIST_02)

    @property
    def list_03(self) -> BulletList:
        """Gets List style 03 (🗹 [checkbox like])"""
        return BulletList(StyleListKind.LIST_03)

    @property
    def list_04(self) -> BulletList:
        """Gets List style 0 (triangle like)"""
        return BulletList(StyleListKind.LIST_04)

    @property
    def list_05(self) -> BulletList:
        """Gets List style 05 (Bullet ꭗ)"""
        return BulletList(StyleListKind.LIST_05)

    @property
    def num_123(self) -> BulletList:
        """Gets List style Numbering ``123``"""
        return BulletList(StyleListKind.NUM_123)

    @property
    def num_abc(self) -> BulletList:
        """Gets List style Numbering ``abc`` (lower case)"""
        return BulletList(StyleListKind.NUM_abc)

    @property
    def num_ABC(self) -> BulletList:
        """Gets List style Numbering ``ABC`` (upper case)"""
        return BulletList(StyleListKind.NUM_ABC)

    @property
    def num_ivx(self) -> BulletList:
        """Gets List style Numbering ``ivx`` (lower case)"""
        return BulletList(StyleListKind.NUM_ivx)

    @property
    def num_IVX(self) -> BulletList:
        """Gets List style Numbering ``IVX`` (upper case)"""
        return BulletList(StyleListKind.NUM_IVX)

    # endregion Style Properties

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.STYLE | FormatKind.STATIC
        return self._format_kind_prop

    @property
    def prop_name(self) -> str:
        """Gets/Sets Character style name"""
        return self._get(self._get_property_name())

    @prop_name.setter
    def prop_name(self, value: StyleListKind | str) -> None:
        if self is BulletList.default:
            raise ValueError("Setting StyleList.default properties is not allowed.")
        self._set(self._get_property_name(), str(value))

    @static_prop
    def default() -> BulletList:  # type: ignore[misc]
        """Gets ``StyleList`` default. Static Property."""
        try:
            return BulletList._DEFAULT  # type: ignore[return-value]
        except AttributeError:
            # empty string name will result in No List, list style being applied
            BulletList._DEFAULT = BulletList(name="")  # type: ignore[assignment]
            BulletList._DEFAULT._is_default_inst = True  # type: ignore[assignment]
        return BulletList._DEFAULT  # type: ignore[return-value]
