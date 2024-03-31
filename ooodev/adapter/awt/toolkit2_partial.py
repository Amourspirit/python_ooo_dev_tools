from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XToolkit2
from ooodev.adapter.awt.toolkit_partial import ToolkitPartial
from ooodev.adapter.awt.data_transfer_provider_access_partial import DataTransferProviderAccessPartial
from ooodev.adapter.awt.system_child_factory_partial import SystemChildFactoryPartial
from ooodev.adapter.awt.message_box_factory_partial import MessageBoxFactoryPartial
from ooodev.adapter.awt.extended_toolkit_partial import ExtendedToolkitPartial


if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
else:
    UnoInterface = Any


class Toolkit2Partial(
    ToolkitPartial,
    DataTransferProviderAccessPartial,
    SystemChildFactoryPartial,
    MessageBoxFactoryPartial,
    ExtendedToolkitPartial,
):
    """
    Partial Class for XToolkit2.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XToolkit2, interface: UnoInterface | None = XToolkit2) -> None:
        """
        Constructor

        Args:
            component (XToolkit2): UNO Component that implements ``com.sun.star.awt.XToolkit2`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XToolkit2``.
        """
        ToolkitPartial.__init__(self, component=component, interface=interface)
        DataTransferProviderAccessPartial.__init__(self, component=component, interface=interface)
        SystemChildFactoryPartial.__init__(self, component=component, interface=interface)
        MessageBoxFactoryPartial.__init__(self, component=component, interface=interface)
        ExtendedToolkitPartial.__init__(self, component=component, interface=interface)
