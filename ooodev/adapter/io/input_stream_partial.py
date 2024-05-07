from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno
from com.sun.star.io import XInputStream

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class InputStreamPartial:
    """
    Partial Class XInputStream.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XInputStream, interface: UnoInterface | None = XInputStream) -> None:
        """
        Constructor

        Args:
            component (XInputStream): UNO Component that implements ``com.sun.star.io.XInputStream`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XInputStream``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XInputStream
    def available(self) -> int:
        """
        states how many bytes can be read or skipped without blocking.

        Note: This method offers no information on whether the EOF has been reached.

        Raises:
            com.sun.star.io.NotConnectedException: ``NotConnectedException``
            com.sun.star.io.IOException: ``IOException``
        """
        return self.__component.available()

    def close_input(self) -> None:
        """
        closes the stream.

        Users must close the stream explicitly when no further reading should be done. (There may exist ring references to chained objects that can only be released during this call. Thus not calling this method would result in a leak of memory or external resources.)

        Raises:
            com.sun.star.io.NotConnectedException: ``NotConnectedException``
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.closeInput()

    def read_bytes(self, num_bytes: int) -> Tuple[int, Tuple[int]]:
        """
        Reads the specified number of bytes in the given sequence.

        The return value specifies the number of bytes which have been put into the sequence. A difference between nBytesToRead and the return value indicates that EOF has been reached. This means that the method blocks until the specified number of bytes are available or the EOF is reached.

        Args:
            num_bytes (int): The number of bytes to read.

        Raises:
            com.sun.star.io.NotConnectedException: ``NotConnectedException``
            com.sun.star.io.BufferSizeExceededException: ``BufferSizeExceededException``
            com.sun.star.io.IOException: ``IOException``
        """
        return self.__component.readBytes((), num_bytes)  # type: ignore

    def read_some_bytes(self, max_read: int) -> Tuple[int, Tuple[int]]:
        """
        Reads the available number of bytes, at maximum nMaxBytesToRead.

        This method is very similar to the readBytes method, except that it has different blocking behaviour.
        The method blocks as long as at least 1 byte is available or EOF has been reached. EOF has only been reached, when the method returns ``0`` and the corresponding byte sequence is empty.
        Otherwise, after the call, aData contains the available, but no more than nMaxBytesToRead, bytes.

        Args:
            max_read (int): The maximum number of bytes to read.

        Raises:
            com.sun.star.io.NotConnectedException: ``NotConnectedException``
            com.sun.star.io.BufferSizeExceededException: ``BufferSizeExceededException``
            com.sun.star.io.IOException: ``IOException``
        """
        return self.__component.readSomeBytes((), max_read)  # type: ignore

    def skip_bytes(self, skip_num: int) -> None:
        """
        Skips the next skip_num bytes (must be positive).

        It is up to the implementation whether this method is blocking the thread or not.

        Raises:
            com.sun.star.io.NotConnectedException: ``NotConnectedException``
            com.sun.star.io.BufferSizeExceededException: ``BufferSizeExceededException``
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.skipBytes(skip_num)

    # endregion XInputStream
