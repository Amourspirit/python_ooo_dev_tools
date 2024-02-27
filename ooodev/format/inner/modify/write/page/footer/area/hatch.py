from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooo.dyn.drawing.hatch_style import HatchStyle

from ooodev.format.inner.common.props.area_hatch_props import AreaHatchProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.modify.write.page.header.area.hatch import Hatch as HeaderHatch
from ooodev.format.writer.style.page.kind.writer_style_page_kind import WriterStylePageKind
from ooodev.units.angle import Angle
from ooodev.utils.color import Color, StandardColor

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT


class Hatch(HeaderHatch):
    """
    Page Footer Hatch

    .. seealso::

        - :ref:`help_writer_format_modify_page_footer_area`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        style: HatchStyle = HatchStyle.SINGLE,
        color: Color = StandardColor.BLACK,
        space: float | UnitT = 0.0,
        angle: Angle | int = 0,
        bg_color: Color = StandardColor.AUTO_COLOR,
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            style (HatchStyle, optional): Specifies the kind of lines used to draw this hatch.
                Default ``HatchStyle.SINGLE``.
            color (:py:data:`~.utils.color.Color`, optional): Specifies the color of the hatch lines. Default ``0``.
            space (float, UnitT, optional): Specifies the space between the lines in the hatch (in ``mm`` units)
                or :ref:`proto_unit_obj`. Default ``0.0``
            angle (Angle, int, optional): Specifies angle of the hatch in degrees.
                Default to ``0``.
            bg_color(:py:data:`~.utils.color.Color`, optional): Specifies the background Color. Set this ``-1`` (default) for no background color.
            style_name (StyleParaKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_page_footer_area`
        """
        super().__init__(
            style=style,
            color=color,
            space=space,
            angle=angle,
            bg_color=bg_color,
            style_name=style_name,
            style_family=style_family,
        )

    # region internal methods
    def _get_inner_props(self) -> AreaHatchProps:
        return AreaHatchProps(
            color="FooterFillColor",
            style="FooterFillStyle",
            bg="FooterFillBackground",
            hatch_prop="FooterFillHatch",
        )

    # endregion internal methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE | FormatKind.FOOTER
        return self._format_kind_prop

    # endregion properties
