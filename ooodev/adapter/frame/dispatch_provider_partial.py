from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple

import uno
from com.sun.star.frame import XDispatchProvider

# com.sun.star.frame.FrameSearchFlag
from ooo.dyn.frame.frame_search_flag import FrameSearchFlagEnum

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.frame import XDispatch
    from com.sun.star.util import URL  # Struct
    from com.sun.star.frame import DispatchDescriptor  # Struct
    from ooodev.utils.type_var import UnoInterface


class DispatchProviderPartial:
    """
    Partial class for XDispatchProvider.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XDispatchProvider, interface: UnoInterface | None = XDispatchProvider) -> None:
        """
        Constructor

        Args:
            component (XDispatchProvider ): UNO Component that implements ``com.sun.star.frame.XDispatchProvider`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDispatchProvider``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XDispatchProvider
    def query_dispatch(
        self, url: URL, target_frame_name: str, search_flags: int | FrameSearchFlagEnum = FrameSearchFlagEnum.AUTO
    ) -> XDispatch | None:
        """
        Returns a dispatch object for the specified URL.

        Args:
            url (URL): Specifies the feature which should be supported by returned dispatch object.
            target_frame_name (str): Specifies the frame which should be the target for this request.
            search_flags (int, FrameSearchFlagEnum, optional): Optional search parameter for finding the frame
                if no special TargetFrameName was used.

        Returns:
            XDispatch: the dispatch object which provides queried functionality
            or None if no dispatch object is available.
        """
        return self.__component.queryDispatch(url, target_frame_name, FrameSearchFlagEnum(search_flags).value)

    def query_dispatches(self, requests: Tuple[DispatchDescriptor, ...]) -> Tuple[XDispatch | None, ...]:
        """
        Returns a list of dispatch objects for the specified URLs.

        Actually this method is redundant to ``query_dispatch()`` to avoid multiple remote calls.

        Args:
            requests (Tuple[DispatchDescriptor, ...]): Tuple of dispatch requests

        Returns:
            Tuple[XDispatch | None, ...]: multiple dispatch interfaces for the specified descriptors at once
        """
        return self.__component.queryDispatches(requests)

    # endregion XDispatchProvider
