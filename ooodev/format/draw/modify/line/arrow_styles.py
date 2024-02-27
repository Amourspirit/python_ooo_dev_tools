"""
Modify Draw Style Arrows.

.. versionadded:: 0.17.13
"""

from __future__ import annotations
from typing import cast, Any, TYPE_CHECKING
import uno

from ooodev.format.draw.style.kind import DrawStyleFamilyKind
from ooodev.format.draw.style.lookup import FamilyGraphics
from ooodev.format.inner.modify.draw.line_properties_style_base_multi import LinePropertiesStyleBaseMulti
from ooodev.format.inner.direct.draw.shape.line.arrow_styles import ArrowStyles as InnerLineArrowStyles

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT
    from ooodev.utils.kind.graphic_arrow_style_kind import GraphicArrowStyleKind


class ArrowStyles(LinePropertiesStyleBaseMulti):
    """
    Arrow Style values.

    .. seealso::

        - :ref:`help_draw_format_modify_shadow_shadow`

    .. versionadded:: 0.17.13
    """

    def __init__(
        self,
        *,
        start_line_name: GraphicArrowStyleKind | str | None = None,
        start_line_width: float | UnitT | None = None,
        start_line_center: bool | None = None,
        end_line_name: GraphicArrowStyleKind | str | None = None,
        end_line_width: float | UnitT | None = None,
        end_line_center: bool | None = None,
        style_name: str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> None:
        """
        Constructor

        Args:
            start_line_name (GraphicArrowStyleKind, str, optional): Start line name. Defaults to ``None``.
            start_line_width (float, UnitT, optional): Start line width in mm units. Defaults to ``None``.
            start_line_center (bool, optional): Start line center. Defaults to ``None``.
            end_line_name (GraphicArrowStyleKind, str, optional): End line name. Defaults to ``None``.
            end_line_width (float, UnitT, optional): End line width in mm units. Defaults to ``None``.
            end_line_center (bool, optional): End line center. Defaults to ``None``.
            style_name (FamilyGraphics, str, optional): Specifies the Style that instance applies to.
                Default is Default ``standard`` Style.
            style_family (str, DrawStyleFamilyKind, optional): Family Style. Defaults to ``graphics``.

        Returns:
            None:

        See Also:
            - :ref:`help_draw_format_modify_shadow_shadow`
        """

        direct = InnerLineArrowStyles(
            start_line_name=start_line_name,
            start_line_width=start_line_width,
            start_line_center=start_line_center,
            end_line_name=end_line_name,
            end_line_width=end_line_width,
            end_line_center=end_line_center,
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = str(style_family)
        self._set_style("direct", direct)

    @classmethod
    def from_style(
        cls,
        doc: Any,
        style_name: str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> ArrowStyles:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (FamilyGraphics, str, optional): Specifies the Style that instance applies to.
                Default is ``FamilyGraphics.DEFAULT_DRAWING_STYLE``.
            style_family (DrawStyleFamilyKind, str, optional): Style family. Default ``DrawStyleFamilyKind.GRAPHICS``.

        Returns:
            ArrowStyles: ``ArrowStyles`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerLineArrowStyles.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerLineArrowStyles:
        """Gets/Sets Inner Font instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerLineArrowStyles, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerLineArrowStyles) -> None:
        if not isinstance(value, InnerLineArrowStyles):
            raise TypeError(f'Expected type of InnerLineArrowStyles, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
