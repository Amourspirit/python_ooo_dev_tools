from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.io import XActiveDataSource

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.io import XOutputStream
    from ooodev.utils.type_var import UnoInterface


class ActiveDataSourcePartial:
    """
    Partial Class XActiveDataSource.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XActiveDataSource, interface: UnoInterface | None = XActiveDataSource) -> None:
        """
        Constructor

        Args:
            component (XActiveDataSource): UNO Component that implements ``com.sun.star.io.XActiveDataSource`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XActiveDataSource``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XActiveDataSource
    def get_output_stream(self) -> XOutputStream:
        """
        Gets the plugged stream.
        """
        return self.__component.getOutputStream()

    def set_output_stream(self, stream: XOutputStream) -> None:
        """
        plugs the output stream.

        If XConnectable is also implemented, this method should query aStream for a XConnectable and connect both.
        """
        self.__component.setOutputStream(stream)

    # endregion XActiveDataSource
