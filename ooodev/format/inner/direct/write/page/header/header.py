# region Import
from __future__ import annotations
import uno
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.common.abstract.abstract_hf import AbstractHF
from ooodev.format.inner.common.props.hf_props import HfProps

# endregion Import


class Header(AbstractHF):
    """
    Page Header Settings

    .. versionadded:: 0.9.2
    """

    @property
    def _props(self) -> HfProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = HfProps(
                on="HeaderIsOn",
                shared="HeaderIsShared",
                shared_first="FirstIsShared",
                margin_left="HeaderLeftMargin",
                margin_right="HeaderRightMargin",
                spacing="HeaderBodyDistance",
                spacing_dyn="HeaderDynamicSpacing",
                height="HeaderHeight",
                height_auto="HeaderIsDynamicHeight",
            )
        return self._props_internal_attributes

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.HEADER
        return self._format_kind_prop
