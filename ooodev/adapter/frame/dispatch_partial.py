from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.frame import XDispatch

# com.sun.star.frame.FrameSearchFlag

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.frame import XStatusListener
    from com.sun.star.beans import PropertyValue  # Struct
    from com.sun.star.util import URL  # Struct
    from ooodev.utils.type_var import UnoInterface


class DispatchPartial:
    """
    Partial class for XDispatch.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XDispatch, interface: UnoInterface | None = XDispatch) -> None:
        """
        Constructor

        Args:
            component (XDispatch ): UNO Component that implements ``com.sun.star.frame.XDispatch`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDispatch``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XDispatch
    def add_status_listener(self, control: XStatusListener, url: URL) -> None:
        """
        Registers a listener of a control for a specific URL at this object to receive status events.

        It is only allowed to register URLs for which this XDispatch was explicitly queried.
        Additional arguments (``#...`` or ``?...``) will be ignored.

        Note: Notifications can't be guaranteed! This will be a part of interface XNotifyingDispatch.
        """
        self.__component.addStatusListener(control, url)

    def dispatch(self, url: URL, *args: PropertyValue) -> None:
        """
        dispatches (executes) a URL

        It is only allowed to dispatch URLs for which this XDispatch was explicitly queried.
        Additional arguments (``#...`` or ``?...``) are allowed.

        Controlling synchronous or asynchronous mode happens via readonly boolean Flag SynchronMode

        By default, and absent any arguments, ``SynchronMode`` is considered ``False`` and the execution is performed asynchronously
        (i.e. dispatch() returns immediately, and the action is performed in the background).
        But when set to ``True``, dispatch() processes the request synchronously
        """
        self.__component.dispatch(url, args)

    def remove_status_listener(self, Control: XStatusListener, url: URL) -> None:
        """
        unregisters a listener from a control.
        """
        self.__component.removeStatusListener(Control, url)

    # endregion XDispatch
