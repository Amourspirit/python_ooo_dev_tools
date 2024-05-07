from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.io import XOutputStream

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class OutputStreamPartial:
    """
    Partial Class XOutputStream.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XOutputStream, interface: UnoInterface | None = XOutputStream) -> None:
        """
        Constructor

        Args:
            component (XOutputStream): UNO Component that implements ``com.sun.star.io.XOutputStream`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XOutputStream``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XOutputStream
    def close_output(self) -> None:
        """
        Gets called to indicate that all data has been written.

        If this method has not yet been called, no attached XInputStream receives an EOF signal. No further bytes may be written after this method has been called.

        Raises:
            com.sun.star.io.NotConnectedException: ``NotConnectedException``
            com.sun.star.io.BufferSizeExceededException: ``BufferSizeExceededException``
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.closeOutput()

    def flush(self) -> None:
        """
        flushes out of the stream any data that may exist in buffers.

        The semantics of this method are rather vague. See com.sun.star.io.XAsyncOutputMonitor.waitForCompletion() for a similar method with very specific semantics, that is useful in certain scenarios.

        Raises:
            com.sun.star.io.NotConnectedException: ``NotConnectedException``
            com.sun.star.io.BufferSizeExceededException: ``BufferSizeExceededException``
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.flush()

    def write_bytes(self, *data: int) -> None:
        """
        Writes the whole sequence to the stream.

        (blocking call)

        Raises:
            com.sun.star.io.NotConnectedException: ``NotConnectedException``
            com.sun.star.io.BufferSizeExceededException: ``BufferSizeExceededException``
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.writeBytes(data)  # type: ignore

    # endregion XOutputStream
