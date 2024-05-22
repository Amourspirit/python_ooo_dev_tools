from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.io import XStream

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.io import XInputStream
    from com.sun.star.io import XOutputStream
    from ooodev.utils.type_var import UnoInterface


class StreamPartial:
    """
    Partial Class XStream.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XStream, interface: UnoInterface | None = XStream) -> None:
        """
        Constructor

        Args:
            component (XStream): UNO Component that implements ``com.sun.star.io.XStream`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XStream``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XStream
    def get_input_stream(self) -> XInputStream:
        """
        Gets the input stream.
        """
        return self.__component.getInputStream()

    def get_output_stream(self) -> XOutputStream:
        """
        Gets the output stream.
        """
        return self.__component.getOutputStream()

    # endregion XStream
