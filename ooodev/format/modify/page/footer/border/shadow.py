from __future__ import annotations
import uno
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation

from .....writer.style.page.kind.style_page_kind import StylePageKind as StylePageKind
from ...header.border.shadow import Shadow as HeaderShadow


class Shadow(HeaderShadow):
    """
    Page Style Footer Border Shadow

    .. versionadded:: 0.9.0
    """

    # region Internal Methods
    def _get_inner_prop_name(self) -> str:
        return "FooterShadowFormat"

    # endregion Internal Methods
