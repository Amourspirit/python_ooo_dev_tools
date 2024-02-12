from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.frame import XComponentLoader

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.type_var import UnoInterface

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue  # struct
    from com.sun.star.lang import XComponent


class ComponentLoaderPartial:
    """
    Partial class for XComponentLoader.
    """

    def __init__(self, component: XComponentLoader, interface: UnoInterface | None = XComponentLoader) -> None:
        """
        Constructor

        Args:
            component (XComponentLoader): UNO Component that implements ``com.sun.star.frame.XComponentLoader`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XComponentLoader``.
        """
        self.__interface = interface
        self.__validate(component)
        self.__component = component

    def __validate(self, component: Any) -> None:
        """
        Validates the component.

        Args:
            component (Any): The component to be validated.
        """
        if self.__interface is None:
            return
        if not mLo.Lo.is_uno_interfaces(component, self.__interface):
            raise mEx.MissingInterfaceError(self.__interface)

    # region XComponentLoader
    def load_component_from_url(
        self, url: str, target_frame_name: str, search_flags: int, args: Tuple[PropertyValue, ...]
    ) -> XComponent:
        """
        Loads a component specified by a URL into the specified new or existing frame.

        To create new documents, use ``private:factory/scalc``, ``private:factory/swriter``, etc.
        Other special protocols (e.g. ``slot:``, ``.uno``) are not allowed and raise a ``com.sun.star.lang.IllegalArgumentException``.

        If a frame with the specified name already exists, it is used, otherwise it is created.
        There exist some special targets which never can be used as real frame names:

        Flags are optional ones and will be used for non special target names only.

        For example, ``ReadOnly`` with a boolean value specifies whether the document is opened read-only.
        ``FilterName`` specifies the component type to create and the filter to use, for example: ``Text - CSV``.
        For more information see ``com.sun.star.document.MediaDescriptor``.

        This interface is a generic one and can be used to start further requests on loaded document or control the lifetime of it (means dispose() it after using).
        The real document service behind this interface can be one of follow three ones:

        Args:
            url (str): Specifies the URL of the document to load.
            target_frame_name (str): Specifies the name of the frame to view the document in.
            search_flags (int): Use the values of ``FrameSearchFlag`` to specify how to find the specified TargetFrameName
            args (Tuple[PropertyValue, ...]): The arguments.

        Raises:
            com.sun.star.io.IOException: ``IOException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``


        Returns:
            XComponent: The loaded component.

        Note:
            ``target_frame_name`` specifies the name of the frame to view the document in.

            - ``_blank``: Always creates a new frame.
            - ``_default``: Special UI functionality/
            - ``_self``: Always uses the same frame.
            - ``_parent``: Address direct parent of frame
            - ``_top``: Indicates top frame of current path in tree.
            - ``_beamer``: Means special sub frame.


        See Also:
            - `FrameSearchFlag <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1frame_1_1FrameSearchFlag.html>`__
            - `XComponentLoader <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1frame_1_1XComponentLoader.html>`__
        """
        return self.__component.loadComponentFromURL(url, target_frame_name, search_flags, args)

    # endregion XComponentLoader
