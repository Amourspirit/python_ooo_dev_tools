from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from ooodev.utils.gen_util import NULL_OBJ
from ooodev.utils.helper.dot_dict import DotDict
from ooodev.io.json.json_custom_props import JsonCustomProps

if TYPE_CHECKING:
    from ooodev.proto.office_document_t import OfficeDocumentT


class JsonCustomPropsPartial:
    """
    Class to add custom properties to a document. The properties are stored in a json file embedded in the document.

    Allows custom properties to be added to a document.

    Note:
        Any value that can be serialized to JSON can be stored as a custom property.
        Classes can implement the :py:class:`ooodev.io.json.json_encoder.JsonEncoder` class to provide custom serialization by overriding the ``on_json_encode()`` method.
    """

    def __init__(self, doc: OfficeDocumentT):
        """
        Constructor.

        Args:
            forms (Any): The component that implements ``XForm``.
            form_name (str, optional): The name to assign to the form. Defaults to "CustomProperties".
            ctl_name (str, optional): The name to assign to the hidden control. Defaults to "CustomProperties".
        """
        self.__jp = JsonCustomProps(doc)

    def get_custom_property(self, name: str, default: Any = NULL_OBJ) -> Any:
        """
        Gets a custom property.

        Args:
            name (str): The name of the property.
            default (Any, optional): The default value to return if the property does not exist.

        Raises:
            AttributeError: If the property is not found.

        Returns:
            Any: The value of the property.
        """
        return self.__jp.get_custom_property(name, default)

    def set_custom_property(self, name: str, value: Any):
        """
        Sets a custom property.

        Args:
            name (str): The name of the property.
            value (Any): The value of the property.

        Raises:
            AttributeError: If the property is a forbidden key.
        """
        self.__jp.set_custom_property(name, value)

    def get_custom_properties(self) -> DotDict:
        """
        Gets custom properties.

        Returns:
            DotDict: custom properties.

        Hint:
            DotDict is a class that allows you to access dictionary keys as attributes or keys.
            DotDict can be imported from ``ooodev.utils.helper.dot_dict.DotDict``.
        """
        return self.__jp.get_custom_properties()

    def set_custom_properties(self, properties: DotDict) -> None:
        """
        Sets custom properties.

        Args:
            properties (DotDict): custom properties to set.

        Hint:
            DotDict is a class that allows you to access dictionary keys as attributes or keys.
            DotDict can be imported from ``ooodev.utils.helper.dot_dict.DotDict``.
        """
        self.__jp.set_custom_properties(properties)

    def remove_custom_property(self, name: str) -> None:
        """
        Removes a custom property.

        Args:
            name (str): The name of the property to remove.

        Raises:
            AttributeError: If the property is a forbidden key.

        Returns:
            None:
        """
        self.__jp.remove_custom_property(name)

    def has_custom_property(self, name: str) -> bool:
        """
        Gets if a custom property exists.

        Args:
            name (str): The name of the property to check.

        Returns:
            bool: ``True`` if the property exists, otherwise ``False``.
        """
        return self.__jp.has_custom_property(name)
