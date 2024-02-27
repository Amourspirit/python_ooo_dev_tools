# region Import
from __future__ import annotations
import uno
from ooo.dyn.table.shadow_location import ShadowLocation

from ooodev.format.calc.style.page.kind.calc_style_page_kind import CalcStylePageKind
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.units.unit_obj import UnitT
from ooodev.utils.color import StandardColor, Color
from ooodev.format.inner.modify.calc.page.header.border.shadow import Shadow as HeaderShadow

# endregion Import


class Shadow(HeaderShadow):
    """
    Page Style Footer Border Shadow

    .. seealso::

        - :ref:`help_calc_format_modify_page_footer_borders`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        location: ShadowLocation = ShadowLocation.BOTTOM_RIGHT,
        color: Color = StandardColor.GRAY,
        transparent: bool = False,
        width: float | UnitT = 1.76,
        style_name: CalcStylePageKind | str = CalcStylePageKind.DEFAULT,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            location (ShadowLocation, optional): contains the location of the shadow.
                Default to ``ShadowLocation.BOTTOM_RIGHT``.
            color (:py:data:`~.utils.color.Color`, optional):contains the color value of the shadow. Defaults to ``StandardColor.GRAY``.
            transparent (bool, optional): Shadow transparency. Defaults to False.
            width (float, UnitT, optional): contains the size of the shadow (in ``mm`` units)
                or :ref:`proto_unit_obj`. Defaults to ``1.76``.
            style_name (CalcStylePageKind, str, optional): Specifies the Page Style that instance applies to.
                Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_modify_page_footer_borders`
        """
        super().__init__(
            location=location,
            color=color,
            transparent=transparent,
            width=width,
            style_name=style_name,
            style_family=style_family,
        )

    # region overrides
    def _get_inner_prop_name(self) -> str:
        return "FooterShadowFormat"

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.DOC | FormatKind.STYLE | FormatKind.FOOTER
        return self._format_kind_prop

    # endregion overrides
