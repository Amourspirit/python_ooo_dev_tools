# region Import
from __future__ import annotations
import uno
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation

from ooodev.format.calc.style.page.kind import CalcStylePageKind as CalcStylePageKind
from ooodev.format.inner.kind.format_kind import FormatKind
from ...header.border.shadow import Shadow as HeaderShadow

# endregion Import


class Shadow(HeaderShadow):
    """
    Page Style Footer Border Shadow

    .. versionadded:: 0.9.0
    """

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
