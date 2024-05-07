from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.io import XActiveDataSink

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.io import XInputStream
    from ooodev.utils.type_var import UnoInterface


class ActiveDataSinkPartial:
    """
    Partial Class XActiveDataSink.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XActiveDataSink, interface: UnoInterface | None = XActiveDataSink) -> None:
        """
        Constructor

        Args:
            component (XActiveDataSink): UNO Component that implements ``com.sun.star.io.XActiveDataSink`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XActiveDataSink``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XActiveDataSink
    def get_input_stream(self) -> XInputStream:
        """
        Gets the plugged stream.
        """
        return self.__component.getInputStream()

    def set_input_stream(self, stream: XInputStream) -> None:
        """
        Plugs the input stream.

        If ``XConnectable`` is also implemented, this method should query aStream for an ``XConnectable`` and connect both.
        """
        self.__component.setInputStream(stream)

    # endregion XActiveDataSink
