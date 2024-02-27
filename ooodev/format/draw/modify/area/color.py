"""Module for Draw Style Fill Coloring."""

# region Import
from __future__ import annotations
from typing import Any, cast
import uno
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind as StyleParaKind
from ooodev.utils import color as mColor
from ooodev.format.inner.direct.write.para.area.color import Color as InnerColor
from ooodev.format.inner.modify.write.fill.fill_style_base_multi import FillStyleBaseMulti
from ooodev.format.draw.style.kind import DrawStyleFamilyKind
from ooodev.format.draw.style.lookup import FamilyGraphics

# endregion Import


class Color(FillStyleBaseMulti):
    """
    Draw Style Fill Coloring

    .. seealso::

        - :ref:`help_draw_format_modify_area_color`

    .. versionadded:: 0.17.9
    """

    def __init__(
        self,
        *,
        color: mColor.Color = mColor.StandardColor.AUTO_COLOR,
        style_name: str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> None:
        """
        Constructor

        Args:
            color (:py:data:`~.utils.color.Color`, optional): Fill Color.
            style_name (FamilyGraphics, str, optional): Specifies the Style that instance applies to.
                Default is Default ``standard`` Style.
            style_family (str, DrawStyleFamilyKind, optional): Family Style. Defaults to ``graphics``.

        Returns:
            None:

        See Also:
            - :ref:`help_draw_format_modify_area_color`
        """

        direct = InnerColor(color=color)
        super().__init__()
        self._style_family_name = str(style_family)
        self._style_name = str(style_name)
        self._set_style("direct", direct)

    @classmethod
    def from_style(
        cls,
        doc: Any,
        style_name: FamilyGraphics | str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> Color:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (FamilyGraphics, str, optional): Specifies the Style that instance applies to.
                Default is ``FamilyGraphics.DEFAULT_DRAWING_STYLE``.
            style_family (DrawStyleFamilyKind, str, optional): Style family. Default ``DrawStyleFamilyKind.GRAPHICS``.

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
    def prop_style_name(self, value: str | FamilyGraphics):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerColor:
        """Gets Inner Color instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerColor, self._get_style_inst("direct"))
        return self._direct_inner
