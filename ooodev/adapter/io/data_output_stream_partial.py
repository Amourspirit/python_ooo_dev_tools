from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.io import XDataOutputStream

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.io.output_stream_partial import OutputStreamPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class DataOutputStreamPartial(OutputStreamPartial):
    """
    Partial Class XDataOutputStream.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XDataOutputStream, interface: UnoInterface | None = XDataOutputStream) -> None:
        """
        Constructor

        Args:
            component (XDataOutputStream): UNO Component that implements ``com.sun.star.io.XDataOutputStream`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDataOutputStream``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        OutputStreamPartial.__init__(self, component=component, interface=None)
        self.__component = component

    # region XDataOutputStream
    def write_boolean(self, value: bool) -> None:
        """
        Writes a boolean.

        It is an 8-bit value. 0 means FALSE; all other values mean TRUE.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.writeBoolean(value)

    def write_byte(self, value: int) -> None:
        """
        Writes an 8-bit byte.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.writeByte(value)

    def write_char(self, value: str) -> None:
        """
        Writes a 16-bit character.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.writeChar(value)

    def write_double(self, value: float) -> None:
        """
        Writes a 64-bit IEEE double.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.writeDouble(value)

    def write_float(self, value: float) -> None:
        """
        Writes a 32-bit IEEE float.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.writeFloat(value)

    def write_hyper(self, value: int) -> None:
        """
        Writes a 64-bit big endian integer.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.writeHyper(value)

    def write_long(self, value: int) -> None:
        """
        Writes a 32-bit big endian integer.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.writeLong(value)

    def write_short(self, value: int) -> None:
        """
        Writes a 16-bit big endian integer.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.writeShort(value)

    def write_utf(self, value: str) -> None:
        """
        Writes a string in UTF format.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.writeUTF(value)

    # endregion XDataOutputStream
