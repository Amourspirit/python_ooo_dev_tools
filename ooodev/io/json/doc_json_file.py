from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from pathlib import Path
import json
from ooodev.io.sfa.sfa import Sfa
from ooodev.io.log.named_logger import NamedLogger
from ooodev.exceptions import ex as mEx

if TYPE_CHECKING:
    from ooodev.proto.office_document_t import OfficeDocumentT


class DocJsonFile:
    """
    Class that read and write JSON files to and from the document.
    """

    def __init__(self, doc: OfficeDocumentT, root_dir: str = "json"):
        self._log = NamedLogger(self.__class__.__name__)
        self._log.debug(f"Initializing JsonFile: {root_dir}")
        self._doc = doc
        self._root_uri = f"vnd.sun.star.tdoc:/{self._doc.runtime_uid}/{root_dir}"
        self._sfa = Sfa()
        if not self._sfa.exists(self._root_uri):
            self._log.debug(f"Creating folder: {self._root_uri}")
            self._sfa.inst.create_folder(self._root_uri)

    def read_json(self, file_name: str) -> Any:
        """
        Reads a JSON file from the document.

        Args:
            file_name (str): The name of the JSON file.

        Raises:
            FileNotFoundError: If the file does not exist.
            JsonError: If there is an error reading the file.
            JsonLoadError: If there is an error parsing the file.

        Returns:
            dict: The JSON data or None if the file does not exist or there is an error reading the file.
        """
        file_uri = f"{self._root_uri}/{file_name}"
        if not self._sfa.exists(file_uri):
            self._log.error(f"File does not exist: {file_uri}")
            raise FileNotFoundError(f"File does not exist: {file_uri}")
        self._log.debug(f"Reading JSON file: {file_uri}")
        try:
            json_str = self._sfa.read_text_file(file_uri)
        except Exception as e:
            self._log.error(f"Error reading JSON file: {file_uri}", exc_info=True)
            raise mEx.JsonError(f"Error reading JSON file: {file_uri}") from e
        try:
            data = json.loads(json_str)
        except Exception as e:
            self._log.error(f"Error parsing JSON file: {file_uri}", exc_info=True)
            raise mEx.JsonLoadError(f"Error parsing JSON file: {file_uri}") from e
        return data

    def write_json(self, file_name: str, data: Any, mode="w") -> None:
        """
        Writes a JSON file to the document.

        Args:
            file_name (str): The name of the JSON file. The file name will have ``.json`` appended to it if it does not already have it.
            data (Any): The data to write to the file.
            mode: (str, optional): The mode to open the file. Defaults to "w".
                mode ``w`` will overwrite the file.
                mode ``a`` will append to the file.
                mode ``x`` will create a new file and write to it failing if the file already exists

        Raises:
            JsonError: If there is an error writing the file.
            JsonDumpError: If there is an error serializing the data.

        Returns:
            None
        """
        file_pth = Path(file_name)
        if file_pth.suffix != ".json":
            file_name = f"{file_name}.json"
        file_uri = f"{self._root_uri}/{file_name}"

        self._log.debug(f"Writing JSON file: {file_uri}")
        try:
            json_str = json.dumps(data, indent=4)
        except Exception as e:
            self._log.error(f"Error serializing JSON data: {file_uri}", exc_info=True)
            raise mEx.JsonDumpError(f"Error serializing JSON data: {file_uri}") from e
        try:
            self._sfa.write_text_file(file_uri, json_str, mode=mode)
        except Exception as e:
            self._log.error(f"Error writing JSON file: {file_uri}", exc_info=True)
            raise mEx.JsonError(f"Error writing JSON file: {file_uri}") from e
        return None

    def file_exist(self, file_name: str) -> bool:
        """
        Checks if a file exists in the document.

        Args:
            file_name (str): The name of the file.

        Returns:
            bool: True if the file exists, False otherwise.
        """
        if not file_name:
            return False
        file_uri = f"{self._root_uri}/{file_name}"
        return self._sfa.exists(file_uri)
