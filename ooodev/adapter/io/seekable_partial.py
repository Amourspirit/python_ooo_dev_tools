from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.io import XSeekable

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class SeekablePartial:
    """
    Partial Class XSeekable.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XSeekable, interface: UnoInterface | None = XSeekable) -> None:
        """
        Constructor

        Args:
            component (XSeekable): UNO Component that implements ``com.sun.star.io.XSeekable`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XSeekable``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    def __len__(self) -> int:
        """
        Returns the length of the stream.

        Returns:
            int: The length of the stream.
        """
        return self.__component.getLength()

    # region XSeekable
    def get_length(self) -> int:
        """
        Returns the length of the stream.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        return self.__component.getLength()

    def get_position(self) -> int:
        """
        returns the current offset of the stream.

        Raises:
            com.sun.star.io.IOException: ``IOException``
        """
        return self.__component.getPosition()

    def seek(self, location: int) -> None:
        """
        Changes the seek pointer to a new location relative to the beginning of the stream.

        This method changes the seek pointer so subsequent reads and writes can take place at a different location in the stream object. It is an error to seek before the beginning of the stream or after the end of the stream.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.io.IOException: ``IOException``
        """
        self.__component.seek(location)

    # endregion XSeekable
