from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.io import XTempFile

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.io.stream_partial import StreamPartial
from ooodev.adapter.io.seekable_partial import SeekablePartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class TempFilePartial(StreamPartial, SeekablePartial):
    """
    Partial Class XTempFile.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTempFile, interface: UnoInterface | None = XTempFile) -> None:
        """
        Constructor

        Args:
            component (XTempFile): UNO Component that implements ``com.sun.star.io.XTempFile`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTempFile``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        StreamPartial.__init__(self, component=component, interface=None)
        SeekablePartial.__init__(self, component=component, interface=None)
        self.__component = component

    # region XTempFile
    @property
    def remove_file(self) -> bool:
        """
        This attribute controls whether the file will be automatically removed on object destruction.
        """
        return self.__component.RemoveFile

    @remove_file.setter
    def remove_file(self, value: bool) -> None:
        self.__component.RemoveFile = value

    @property
    def resource_name(self) -> str:
        """
        This attribute specifies the temp file name.
        """
        # '/tmp/lu23986334kdxgl.tmp/lu23986334kdxhw.tmp'
        return self.__component.ResourceName

    @property
    def uri(self) -> str:
        """
        This attribute specifies the URL of the temp file.
        """
        # 'file:///tmp/lu23986334kdxgl.tmp/lu23986334kdxhw.tmp'
        return self.__component.Uri

    # endregion XTempFile
