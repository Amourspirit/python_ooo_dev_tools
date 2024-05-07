from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.io import XTextInputStream

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.io.input_stream_partial import InputStreamPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface

# see tests.test_adapter.test_ucb.test_simple_file_access.test_simple_file_access


class TextInputStreamPartial(InputStreamPartial):
    """
    Partial Class XTextInputStream.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextInputStream, interface: UnoInterface | None = XTextInputStream) -> None:
        """
        Constructor

        Args:
            component (XTextInputStream): UNO Component that implements ``com.sun.star.io.XTextInputStream`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTextInputStream``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        InputStreamPartial.__init__(self, component=component, interface=None)
        self.__component = component

    # region XTextInputStream
    def is_eof(self) -> bool:
        """
        Returns the EOF status.

        This method has to be used to detect if the end of the stream is reached.

        Important: This cannot be detected by asking for an empty string because that can be a valid return value of readLine() (if the line is empty) and readString() (if a delimiter is directly followed by the next one).

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        return self.__component.isEOF()

    def read_line(self) -> str:
        """
        reads text until a line break (CR, LF, or CR/LF) or EOF is found and returns it as string (without CR, LF).

        The read characters are converted according to the encoding defined by setEncoding(). If EOF is already reached before calling this method an empty string is returned.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        return self.__component.readLine()

    def read_string(self, remove_delimiter: bool, *delimiters: str) -> str:
        """
        reads text until one of the given delimiter characters or EOF is found and returns it as string (without delimiter).

        Important: ``CR/LF`` is not used as default delimiter!
        So if no delimiter is defined or none of the delimiters is found, the stream will be read to ``EOF``.
        The read characters are converted according to the encoding defined by ``set_encoding()``.
        If ``EOF`` is already reached before calling this method an empty string is returned.

        Args:
            remove_delimiter (bool): If ``True``, the delimiter character is not included in the returned string.
            delimiters (str, optional): One or more delimiter characters.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        return self.__component.readString(delimiters, remove_delimiter)

    def set_encoding(self, encoding: str) -> None:
        """
        sets character encoding.
        """
        self.__component.setEncoding(encoding)

    # endregion XTextInputStream
