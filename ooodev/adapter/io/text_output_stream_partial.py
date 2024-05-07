from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.io import XTextOutputStream

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.io.output_stream_partial import OutputStreamPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface

# see tests.test_adapter.test_ucb.test_simple_file_access.test_simple_file_access


class TextOutputStreamPartial(OutputStreamPartial):
    """
    Partial Class XTextOutputStream.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextOutputStream, interface: UnoInterface | None = XTextOutputStream) -> None:
        """
        Constructor

        Args:
            component (XTextOutputStream): UNO Component that implements ``com.sun.star.io.XTextOutputStream`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTextOutputStream``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        OutputStreamPartial.__init__(self, component=component, interface=None)
        self.__component = component

    # region XTextOutputStream
    def set_encoding(self, encoding: str) -> None:
        """
        Sets character encoding.
        """
        self.__component.setEncoding(encoding)

    def write_string(self, text: str) -> None:
        """
        Writes a string to the stream using the encoding defined by setEncoding().

        Line breaks or delimiters that may be necessary to support XTextInputStream.readLine() and XTextInputStream.readString() have to be added manually to the parameter string.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.writeString(text)

    # endregion XTextOutputStream
