from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.frame import XDesktop2

from ooodev.adapter.frame.dispatch_provider_partial import DispatchProviderPartial
from ooodev.adapter.frame.dispatch_provider_interception_partial import DispatchProviderInterceptionPartial
from ooodev.adapter.frame.frames_supplier_partial import FramesSupplierPartial
from ooodev.adapter.frame.desktop_partial import DesktopPartial
from ooodev.adapter.frame.component_loader_partial import ComponentLoaderPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class Desktop2Partial(
    DispatchProviderPartial,
    DispatchProviderInterceptionPartial,
    FramesSupplierPartial,
    DesktopPartial,
    ComponentLoaderPartial,
):
    """
    Partial class for XDesktop2.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XDesktop2, interface: UnoInterface | None = XDesktop2) -> None:
        """
        Constructor

        Args:
            component (XDesktop2 ): UNO Component that implements ``com.sun.star.frame.XDesktop2`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDesktop2``.
        """

        DispatchProviderPartial.__init__(self, component=component, interface=interface)
        DispatchProviderInterceptionPartial.__init__(self, component=component, interface=interface)
        FramesSupplierPartial.__init__(self, component=component, interface=interface)
        DesktopPartial.__init__(self, component=component, interface=interface)
        ComponentLoaderPartial.__init__(self, component=component, interface=interface)
