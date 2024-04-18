from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter._helper.builder import builder_helper
from ooodev.adapter._helper.builder.comp_defaults_partial import CompDefaultsPartial
from ooodev.adapter.component_prop import ComponentProp
from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.adapter.sheet.cell_range_referrer_partial import CellRangeReferrerPartial
from ooodev.adapter.sheet.view_pane_partial import ViewPanePartial


if TYPE_CHECKING:
    from com.sun.star.sheet import SpreadsheetViewPane  # service


class _SpreadsheetViewPaneComp(ComponentProp):

    def __eq__(self, other: Any) -> bool:
        if not isinstance(other, ComponentProp):
            return False
        if self is other:
            return True
        if self.component is other.component:
            return True
        return self.component == other.component

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.SpreadsheetViewPane",)

    # endregion Overrides


class SpreadsheetViewPaneComp(
    _SpreadsheetViewPaneComp, ViewPanePartial, CellRangeReferrerPartial, CompDefaultsPartial
):
    """
    Class for managing SpreadsheetViewPane Component.

    .. versionadded:: 0.20.0
    """

    # Some implementations may also implement com.sun.star.view.XControlAccess

    # pylint: disable=unused-argument
    def __new__(cls, component: Any, *args, **kwargs):
        builder = get_builder(component=component)
        builder_helper.builder_add_comp_defaults(builder)

        builder_only = kwargs.get("_builder_only", False)
        if builder_only:
            # cast to prevent type checker error
            return cast(Any, builder)
        inst = builder.build_class(
            name="ooodev.adapter.sheet.spreadsheet_view_pane_comp.SpreadsheetViewPaneComp",
            base_class=_SpreadsheetViewPaneComp,
        )
        return inst

    def __init__(self, component: SpreadsheetViewPane) -> None:
        """
        Constructor

        Args:
            component (SpreadsheetViewPane): UNO Volatile Result Component
        """
        # this it not actually called as __new__ is overridden
        pass


def get_builder(component: Any, **kwargs) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    builder = DefaultBuilder(component)
    auto_interface = kwargs.get("auto_interface", True)
    if auto_interface:
        builder.auto_interface()
    return builder
