# region Imports
from __future__ import annotations
from typing import Tuple, cast
import uno
from ooo.dyn.drawing.fill_style import FillStyle

from ooodev.format.calc.style.page.kind.calc_style_page_kind import CalcStylePageKind
from ooodev.format.inner.common.abstract.abstract_fill_color import AbstractColor
from ooodev.format.inner.common.props.fill_color_props import FillColorProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.modify.calc.cell_style_base_multi import CellStyleBaseMulti
from ooodev.utils import color as mColor
from ooodev.utils.color import StandardColor

# endregion Imports


class InnerColor(AbstractColor):
    """
    Page Style Color.

    .. versionadded:: 0.9.0
    """

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.PageProperties", "com.sun.star.style.PageStyle")
        return self._supported_services_values

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PAGE | FormatKind.STYLE
        return self._format_kind_prop

    # region Properties

    @property
    def prop_color(self) -> mColor.Color:
        """Gets/Sets color"""
        return self._get(self._props.color)

    @prop_color.setter
    def prop_color(self, value: mColor.Color):
        if value >= 0:
            self._set(self._props.color, value)
            if self._props.style:
                self._set(self._props.style, FillStyle.SOLID)
            if self._props.bg:
                self._set(self._props.bg, value)
        else:
            self._set(self._props.color, -1)
            if self._props.style:
                self._set(self._props.style, FillStyle.NONE)
            if self._props.bg:
                self._set(self._props.bg, -1)

    @property
    def _props(self) -> FillColorProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = FillColorProps(color="BackColor", style="", bg="BackgroundColor")
        return self._props_internal_attributes

    # endregion Properties


class Color(CellStyleBaseMulti):
    """
    Page Style Color.

    .. seealso::

        - :ref:`help_calc_format_modify_page_background`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        color: mColor.Color = StandardColor.AUTO_COLOR,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            color (:py:data:`~.utils.color.Color`, optional): FillColor Color
            style_name (CalcStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_page_background`
        """

        direct = InnerColor(color=color)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct)

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> Color:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (CalcStylePageKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            Color: ``Color`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerColor.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct)
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | CalcStylePageKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerColor:
        """Gets/Sets Inner Color instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerColor, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerColor) -> None:
        if not isinstance(value, InnerColor):
            raise TypeError(f'Expected type of InnerColor, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value)
