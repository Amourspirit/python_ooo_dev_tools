from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.io import XDataInputStream

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.io.input_stream_partial import InputStreamPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class DataInputStreamPartial(InputStreamPartial):
    """
    Partial Class XDataInputStream.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XDataInputStream, interface: UnoInterface | None = XDataInputStream) -> None:
        """
        Constructor

        Args:
            component (XDataInputStream): UNO Component that implements ``com.sun.star.io.XDataInputStream`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDataInputStream``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        InputStreamPartial.__init__(self, component=component, interface=None)
        self.__component = component

    # region XDataInputStream
    def read_boolean(self) -> int:
        """
        Reads in a boolean.

        It is an 8-bit value. 0 means FALSE; all other values mean TRUE.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        return self.__component.readBoolean()

    def read_byte(self) -> int:
        """
        Reads an 8-bit byte.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        return self.__component.readByte()

    def read_char(self) -> str:
        """
        Reads a 16-bit unicode character.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        return self.__component.readChar()

    def read_double(self) -> float:
        """
        Reads a 64-bit IEEE double.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        return self.__component.readDouble()

    def read_float(self) -> float:
        """
        Reads a 32-bit IEEE float.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        return self.__component.readFloat()

    def read_hyper(self) -> int:
        """
        Reads a 64-bit big endian integer.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        return self.__component.readHyper()

    def read_long(self) -> int:
        """
        Reads a 32-bit big endian integer.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        return self.__component.readLong()

    def read_short(self) -> int:
        """
        Reads a 16-bit big endian integer.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        return self.__component.readShort()

    def read_utf(self) -> str:
        """
        Reads a string of UTF encoded characters.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        return self.__component.readUTF()

    # endregion XDataInputStream
