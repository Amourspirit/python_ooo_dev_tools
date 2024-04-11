from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.util import XURLTransformer

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.util import URL  # struct
    from ooodev.utils.type_var import UnoInterface


class URLTransformerPartial:
    """
    Partial class for XURLTransformer.
    """

    def __init__(self, component: XURLTransformer, interface: UnoInterface | None = XURLTransformer) -> None:
        """
        Constructor

        Args:
            component (XURLTransformer): UNO Component that implements ``com.sun.star.util.XURLTransformer.XURLTransformer`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XURLTransformer``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XURLTransformer

    def assemble(self, url: URL) -> Tuple[bool, URL]:
        """
        Assembles the parts of the URL specified by aURL and stores it into URL.Complete

        Returns:
            Tuple[bool, URL]: The first element is a flag indicating whether the assembling was successful. The second element is the assembled URL.
        """
        # url is a In Out parameter in the IDL
        return self.__component.assemble(url)  # type: ignore

    def get_presentation(self, url: URL, with_password: bool) -> str:
        """
        Gets a representation of the URL for UI purposes only

        Sometimes it can be useful to show a URL on an user interface in a more \"human readable\" form. Such URL can't be used on any API call, but make it easier for the user to understand it.
        """
        return self.__component.getPresentation(url, with_password)

    def parse_smart(self, url: URL, smart_protocol: str) -> Tuple[bool, URL]:
        """
        Parses the string in URL.Complete, which may contain a syntactically complete URL or is specified by the provided protocol

        The implementation can use smart functions to correct or interpret URL.Complete if it is not a syntactically complete URL. The parts of the URL are stored in the other fields of aURL.

        Returns:
            Tuple[bool, URL]: The first element is a flag indicating whether the parsing was successful. The second element is the parsed URL.
        """
        # url is a In Out parameter in the IDL
        return self.__component.parseSmart(url, smart_protocol)  # type: ignore

    def parse_strict(self, url: URL) -> Tuple[bool, URL]:
        """
        Parses the string in URL.Complete which should contain a syntactically complete URL.

        The implementation is allowed to correct minor failures in URL.Complete if the meaning of the URL remain unchanged. Parts of the URL are stored in the other fields of aURL.

        Returns:
            Tuple[bool, URL]: The first element is a flag indicating whether the parsing was successful. The second element is the parsed URL.
        """
        # url is a In Out parameter in the IDL
        return self.__component.parseStrict(url)  # type: ignore

    # endregion XURLTransformer


def get_builder(component: Any) -> Any:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component)
    builder.auto_add_interface("com.sun.star.util.XURLTransformer", optional=False)
    return builder
