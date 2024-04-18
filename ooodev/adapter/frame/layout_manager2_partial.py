from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.frame import XLayoutManager2

from ooodev.utils.builder.default_builder import DefaultBuilder
from ooodev.adapter.frame import layout_manager_partial
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.frame import frame_action_events
from ooodev.adapter.ui import ui_configuration_events
from ooodev.adapter.frame.menu_bar_merging_acceptor_partial import MenuBarMergingAcceptorPartial
from ooodev.adapter.frame.layout_manager_event_broadcaster_partial import LayoutManagerEventBroadcasterPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class LayoutManager2Partial(
    layout_manager_partial.LayoutManagerPartial,
    frame_action_events.FrameActionEvents,
    ui_configuration_events.UIConfigurationEvents,
    MenuBarMergingAcceptorPartial,
    LayoutManagerEventBroadcasterPartial,
):
    """
    Partial class for XLayoutManager2.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XLayoutManager2, interface: UnoInterface | None = XLayoutManager2) -> None:
        """
        Constructor

        Args:
            component (XLayoutManager2 ): UNO Component that implements ``com.sun.star.frame.XLayoutManager2`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XLayoutManager2``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        layout_manager_partial.LayoutManagerPartial.__init__(self, component=component, interface=None)
        MenuBarMergingAcceptorPartial.__init__(self, component=component, interface=None)
        LayoutManagerEventBroadcasterPartial.__init__(self, component=component, interface=None)
        frame_action_events.FrameActionEvents.__init__(self, cb=frame_action_events.on_lazy_cb)
        ui_configuration_events.UIConfigurationEvents.__init__(self, cb=ui_configuration_events.on_lazy_cb)


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """

    builder = DefaultBuilder(component)
    builder.auto_add_interface("com.sun.star.frame.XLayoutManager2", optional=False)
    builder.merge(layout_manager_partial.get_builder(component))
    builder.set_omit(
        "ooodev.adapter.frame.layout_manager_partial.LayoutManagerPartial",
        "ooodev.adapter.frame.frame_action_events.FrameActionEvents",
        "ooodev.adapter.ui.ui_configuration_events.UIConfigurationEvents",
        "ooodev.adapter.frame.menu_bar_merging_acceptor_partial.MenuBarMergingAcceptorPartial",
        "ooodev.adapter.frame.layout_manager_event_broadcaster_partial.LayoutManagerEventBroadcasterPartial",
    )

    return builder
