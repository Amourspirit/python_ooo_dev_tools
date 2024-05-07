from __future__ import annotations
from typing import Any, TYPE_CHECKING

import uno
from com.sun.star.xml.xpath import XXPathAPI

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.xml.dom.node_list_comp import NodeListComp

if TYPE_CHECKING:
    from com.sun.star.xml.xpath import XXPathObject
    from com.sun.star.xml.dom import XNode
    from com.sun.star.xml.dom import XNodeList
    from com.sun.star.xml.xpath import XXPathExtension
    from ooodev.utils.type_var import UnoInterface


class XPathAPIPartial:
    """
    Partial class for XXPathAPI.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XXPathAPI, interface: UnoInterface | None = XXPathAPI) -> None:
        """
        Constructor

        Args:
            component (XXPathAPI ): UNO Component that implements ``com.sun.star.xml.xpath.XXPathAPI`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XXPathAPI``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XXPathAPI
    def eval(self, context_node: XNode, expr: str) -> XXPathObject:
        """
        Evaluate XPath Expression.

        Raises:
            com.sun.star.xml.xpath.XPathException: ``XPathException``
        """
        return self.__component.eval(context_node, expr)

    def eval_ns(self, context_node: XNode, expr: str, ns_node: XNode) -> XXPathObject:
        """
        Evaluate XPath Expression.

        Raises:
            com.sun.star.xml.xpath.XPathException: ``XPathException``
        """
        return self.__component.evalNS(context_node, expr, ns_node)

    def register_extension(self, service_name: str) -> None:
        """
        Register Extension
        """
        self.__component.registerExtension(service_name)

    def register_extension_instance(self, ext: XXPathExtension) -> None:
        """
        Register Extension Instance
        """
        self.__component.registerExtensionInstance(ext)

    def register_ns(self, prefix: str, url: str) -> None:
        """
        Register Namespace
        """
        self.__component.registerNS(prefix, url)

    def select_node_list(self, context_node: XNode, expr: str) -> NodeListComp:
        """
        Evaluate an XPath expression to select a list of nodes.

        Raises:
            com.sun.star.xml.xpath.XPathException: ``XPathException``
        """
        return NodeListComp(self.__component.selectNodeList(context_node, expr))

    def select_node_list_ns(self, context_node: XNode, expr: str, ns_node: XNode) -> NodeListComp:
        """
        Evaluate an XPath expression to select a list of nodes.

        Raises:
            com.sun.star.xml.xpath.XPathException: ``XPathException``
        """
        return NodeListComp(self.__component.selectNodeListNS(context_node, expr, ns_node))

    def select_single_node(self, context_node: XNode, expr: str) -> XNode:
        """
        Evaluate an XPath expression to select a single node.

        Raises:
            com.sun.star.xml.xpath.XPathException: ``XPathException``
        """
        return self.__component.selectSingleNode(context_node, expr)

    def select_single_node_ns(self, context_node: XNode, expr: str, ns_node: XNode) -> XNode:
        """
        Evaluate an XPath expression to select a single node.

        Raises:
            com.sun.star.xml.xpath.XPathException: ``XPathException``
        """
        return self.__component.selectSingleNodeNS(context_node, expr, ns_node)

    def unregister_ns(self, prefix: str, url: str) -> None:
        """
        Un-register namespace.
        """
        self.__component.unregisterNS(prefix, url)

    # endregion XXPathAPI
