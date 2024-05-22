from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno
from com.sun.star.ucb import XContentCreator

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.ucb import XContent
    from com.sun.star.ucb import ContentInfo  # struct
    from ooodev.utils.type_var import UnoInterface


class ContentCreatorPartial:
    """
    Partial Class XContentCreator.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XContentCreator, interface: UnoInterface | None = XContentCreator) -> None:
        """
        Constructor

        Args:
            component (XContentCreator): UNO Component that implements ``com.sun.star.ucb.XContentCreator`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XContentCreator``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XContentCreator
    def create_new_content(self, info: ContentInfo) -> XContent:
        """
        Creates a new content of given type.
        """
        return self.__component.createNewContent(info)

    def query_creatable_contents_info(self) -> Tuple[ContentInfo, ...]:
        """
        Returns a list with information about the creatable contents.
        """
        return self.__component.queryCreatableContentsInfo()

    # endregion XContentCreator
