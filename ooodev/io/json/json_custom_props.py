from __future__ import annotations
from typing import Any, TYPE_CHECKING
import contextlib
from ooodev.utils.gen_util import NULL_OBJ
from ooodev.utils.helper.dot_dict import DotDict
from ooodev.io.json.doc_json_file import DocJsonFile
from ooodev.io.log.named_logger import NamedLogger
from ooodev import get_version


if TYPE_CHECKING:
    from ooodev.proto.office_document_t import OfficeDocumentT


class JsonCustomProps:
    """
    Class to add custom properties to a document. The properties are stored in a json file embedded in the document.

    Allows custom properties to be added to a document.

    Note:
        Any value that can be serialized to JSON can be stored as a custom property.
        Classes can implement the :py:class:`ooodev.io.json.json_encoder.JsonEncoder` class to provide custom serialization by overriding the ``on_json_encode()`` method.
    """

    def __init__(self, doc: OfficeDocumentT, file_name: str = "DocumentCustomProperties.json"):
        """
        Constructor.

        Args:
            doc (Any): The document.
            file_name (str, optional): The name of the file to store the properties. Defaults to "DocumentCustomProperties.json".
        """
        self._log = NamedLogger(self.__class__.__name__)

        self._json_doc = DocJsonFile(doc, "json")
        if not file_name.endswith(".json"):
            file_name = f"{file_name}.json"
        self._name = file_name
        self._props = self._get_custom_properties()
        # please the type checker

    def _get_custom_properties(self) -> dict:
        """
        Loads custom properties from the hidden control.
        """
        if not self._json_doc.file_exist(self._name):
            self._log.debug(f"File does not exist: {self._name}. Creating empty dictionary.")
            return {}
        try:
            result = self._json_doc.read_json(self._name)
            return result.get("data", {})
        except Exception:
            self._log.error(f"Error reading JSON file: {self._name}. Returning empty dictionary.", exc_info=True)
            return {}

    def _save_properties(self, data: dict) -> None:
        """
        Saves custom properties to the hidden control.

        Args:
            properties (dict): The properties to save.
        """
        try:
            json_data = {
                "id": "ooodev.io.json.json_custom_props.JsonCustomProps",
                "version": get_version(),
                "data": data,
            }
            self._json_doc.write_json(self._name, json_data)
        except Exception:
            self._log.error(f"Error writing JSON file: {self._name}", exc_info=True)

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
        result = self._props.get(name, default)
        if result is NULL_OBJ:
            self._log.error(f"Property '{name}' not found.")
            raise AttributeError(f"Property '{name}' not found.")
        return result

    def set_custom_property(self, name: str, value: Any):
        """
        Sets a custom property.

        Args:
            name (str): The name of the property.
            value (Any): The value of the property.

        Raises:
            AttributeError: If the property is a forbidden key.
        """
        with contextlib.suppress(Exception):
            props = self._props.copy()
            props[name] = value
            self._save_properties(props)
            self._props = props

    def get_custom_properties(self) -> DotDict:
        """
        Gets custom properties.

        Returns:
            DotDict: custom properties.

        Hint:
            DotDict is a class that allows you to access dictionary keys as attributes or keys.
            DotDict can be imported from ``ooodev.utils.helper.dot_dict.DotDict``.
        """
        dd = DotDict()
        dd.update(self._props)
        return dd

    def set_custom_properties(self, properties: DotDict) -> None:
        """
        Sets custom properties.

        Args:
            properties (DotDict): custom properties to set.

        Hint:
            DotDict is a class that allows you to access dictionary keys as attributes or keys.
            DotDict can be imported from ``ooodev.utils.helper.dot_dict.DotDict``.
        """
        with contextlib.suppress(Exception):
            props = self._props.copy()
            props.update(properties.copy_dict())
            self._save_properties(props)
            self._props = props

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
        with contextlib.suppress(Exception):
            props = self._props.copy()
            if name in props:
                del props[name]
                self._save_properties(props)
                self._props = props

    def has_custom_property(self, name: str) -> bool:
        """
        Gets if a custom property exists.

        Args:
            name (str): The name of the property to check.

        Returns:
            bool: ``True`` if the property exists, otherwise ``False``.
        """
        return name in self._props
