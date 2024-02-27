# region Import
from __future__ import annotations
import uno
from ooo.dyn.drawing.hatch_style import HatchStyle as HatchStyle

from ooodev.format.inner.common.props.area_hatch_props import AreaHatchProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.preset.preset_hatch import PresetHatchKind as PresetHatchKind
from ooodev.format.inner.direct.write.page.header.area.hatch import Hatch as HeaderHatch

# endregion Import


class Hatch(HeaderHatch):
    """
    Hatch area of the footer.

    .. versionadded:: 0.9.2
    """

    # region Properties
    @property
    def _props(self) -> AreaHatchProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = AreaHatchProps(
                color="FooterFillColor",
                style="FooterFillStyle",
                bg="FooterFillBackground",
                hatch_prop="FooterFillHatch",
            )
        return self._props_internal_attributes

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.FOOTER
        return self._format_kind_prop

    # endregion Properties
