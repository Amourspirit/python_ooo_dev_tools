# region Import
from __future__ import annotations
import uno
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.writer.style.page.kind.writer_style_page_kind import WriterStylePageKind
from ooodev.format.inner.common.abstract.abstract_hf import AbstractHF
from ooodev.format.inner.common.props.hf_props import HfProps

# endregion Import


class Footer(AbstractHF):
    """
    Page Footer Settings

    .. versionadded:: 0.9.2
    """

    @property
    def _props(self) -> HfProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = HfProps(
                on="FooterIsOn",
                shared="FooterIsShared",
                shared_first="FirstIsShared",
                margin_left="FooterLeftMargin",
                margin_right="FooterRightMargin",
                spacing="FooterBodyDistance",
                spacing_dyn="FooterDynamicSpacing",
                height="FooterHeight",
                height_auto="FooterIsDynamicHeight",
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
