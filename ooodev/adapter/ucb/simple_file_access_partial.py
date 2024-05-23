from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import datetime
import uno
from com.sun.star.ucb import XSimpleFileAccess

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.date_time_util import DateUtil

if TYPE_CHECKING:
    from com.sun.star.io import XStream
    from com.sun.star.io import XInputStream
    from com.sun.star.io import XOutputStream
    from com.sun.star.task import XInteractionHandler
    from ooodev.utils.type_var import UnoInterface


class SimpleFileAccessPartial:
    """
    Partial Class XSimpleFileAccess.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XSimpleFileAccess, interface: UnoInterface | None = XSimpleFileAccess) -> None:
        """
        Constructor

        Args:
            component (XSimpleFileAccess): UNO Component that implements ``com.sun.star.ucb.XSimpleFileAccess`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XSimpleFileAccess``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XSimpleFileAccess

    def copy(self, source_url: str, dest_url: str) -> None:
        """
        Copies a file.

        Raises:
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.uno.Exception: ``Exception``
        """
        self.__component.copy(source_url, dest_url)

    def create_folder(self, new_folder_url: str) -> None:
        """
        Creates a new Folder.

        Raises:
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.uno.Exception: ``Exception``
        """
        self.__component.createFolder(new_folder_url)

    def exists(self, file_url: str) -> bool:
        """
        Checks if a file exists.

        Raises:
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.exists(file_url)

    def get_content_type(self, file_url: str) -> str:
        """
        Returns the content type of a file.

        Raises:
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.getContentType(file_url)

    def get_date_time_modified(self, file_url: str) -> datetime.datetime:
        """
        Returns the last modified date for the file.

        Raises:
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.uno.Exception: ``Exception``
        """
        return DateUtil.uno_dt_to_dt(self.__component.getDateTimeModified(file_url))

    def get_folder_contents(self, folder_url: str, include_folders: bool = True) -> Tuple[str, ...]:
        """
        Returns the contents of a folder.

        Args:
            folder_url (str): The URL of the folder.
            include_folders (bool, optional): If ``True``, folders are included in the result. Defaults to ``True``.

        Raises:
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.getFolderContents(folder_url, include_folders)

    def get_size(self, file_url: str) -> int:
        """
        Returns the size of a file.

        Raises:
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.getSize(file_url)

    def is_folder(self, file_url: str) -> bool:
        """
        Checks if a URL represents a folder.

        Raises:
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.isFolder(file_url)

    def is_read_only(self, file_url: str) -> bool:
        """
        Checks if a file is ``read only``.

        Raises:
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.isReadOnly(file_url)

    def kill(self, file_url: str) -> None:
        """
        Removes a file.

        If the URL represents a folder, the folder will be removed, even if it's not empty.

        Raises:
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.uno.Exception: ``Exception``
        """
        self.__component.kill(file_url)

    def move(self, source_url: str, dest_url: str) -> None:
        """
        Moves a file.

        Raises:
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.uno.Exception: ``Exception``
        """
        self.__component.move(source_url, dest_url)

    def open_file_read(self, file_url: str) -> XInputStream:
        """
        Opens file to read.

        Raises:
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.openFileRead(file_url)

    def open_file_read_write(self, file_url: str) -> XStream:
        """
        Opens file to read and write.

        Raises:
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.openFileReadWrite(file_url)

    def open_file_write(self, file_url: str) -> XOutputStream:
        """
        Opens file to write.

        Raises:
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.uno.Exception: ``Exception``
        """
        return self.__component.openFileWrite(file_url)

    def set_interaction_handler(self, handler: XInteractionHandler) -> None:
        """
        Sets an interaction handler to be used for further operations.

        A default interaction handler is available as service com.sun.star.task.InteractionHandler. The documentation of this service also contains further information about the interaction handler concept.
        """
        self.__component.setInteractionHandler(handler)

    def set_read_only(self, file_url: str, read_only: bool) -> None:
        """
        Sets the ``read only`` of a file according to the boolean parameter, if the actual process has the right to do so.

        Raises:
            com.sun.star.ucb.CommandAbortedException: ``CommandAbortedException``
            com.sun.star.uno.Exception: ``Exception``
        """
        self.__component.setReadOnly(file_url, read_only)

    # endregion XSimpleFileAccess
