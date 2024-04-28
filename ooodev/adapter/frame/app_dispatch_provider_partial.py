from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.frame import XAppDispatchProvider


from .dispatch_information_provider_partial import DispatchInformationProviderPartial
from .dispatch_provider_partial import DispatchProviderPartial


if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class AppDispatchProviderPartial(DispatchInformationProviderPartial, DispatchProviderPartial):
    """
    Partial class for XAppDispatchProvider.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XAppDispatchProvider, interface: UnoInterface | None = XAppDispatchProvider) -> None:
        """
        Constructor

        Args:
            component (XAppDispatchProvider ): UNO Component that implements ``com.sun.star.frame.XAppDispatchProvider`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XAppDispatchProvider``.
        """

        DispatchInformationProviderPartial.__init__(self, component=component, interface=interface)
        DispatchProviderPartial.__init__(self, component=component, interface=interface)
