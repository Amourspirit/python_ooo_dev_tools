from __future__ import annotations
from typing import TYPE_CHECKING

import uno
from com.sun.star.frame import XUIControllerFactory

from ooodev.adapter.lang.multi_component_factory_partial import MultiComponentFactoryPartial
from ooodev.adapter.frame.ui_controller_registration_partial import UIControllerRegistrationPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class UIControllerFactoryPartial(MultiComponentFactoryPartial, UIControllerRegistrationPartial):
    """
    Partial class for XUIControllerFactory.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XUIControllerFactory, interface: UnoInterface | None = XUIControllerFactory) -> None:
        """
        Constructor

        Args:
            component (XUIControllerFactory  ): UNO Component that implements ``com.sun.star.frame.XUIControllerFactory`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XUIControllerFactory``.
        """

        MultiComponentFactoryPartial.__init__(self, component=component, interface=interface)
        UIControllerRegistrationPartial.__init__(self, component=component, interface=interface)
