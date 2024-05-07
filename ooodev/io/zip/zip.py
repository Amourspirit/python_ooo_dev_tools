# region Imports
from __future__ import annotations
from typing import Any, cast, Tuple, TYPE_CHECKING
from pathlib import Path
import uno
from ooo.dyn.embed.element_modes import ElementModes

from ooodev.utils.type_var import PathOrStr
from ooodev.adapter.packages.zip.zip_file_access_comp import ZipFileAccessComp
from ooodev.adapter.embed.storage_factory_comp import StorageFactoryComp
from ooodev.utils import props as mProps

if TYPE_CHECKING:
    from com.sun.star.frame import XStorable
# endregion Imports


class ZIP:
    """Class for managing ZIP files."""

    @staticmethod
    def get_zip_content_as_byte_array(doc: Any, fnm: PathOrStr) -> Tuple[int, ...]:
        """
        Gets the content of a file within the document as a byte array.

        Args:
            doc (Any): A UNO Document the supports ``com.sun.star.frame.XStorable`` interface.
            fnm (PathOrStr): Path within the document to the file.

        Raises:
            Exception: If failed to read bytes.

        Returns:
            Tuple[int, ...]: The content of the file as a byte array.

        Warning:
            The ``doc`` must be an existing doc. A newly created doc that has not been saved will not work.

        Example:
            .. code-block:: python

                from ooodev.write import WriteDoc

                doc = WriteDoc.from_current_doc()

                content_xml_bytes = ZIP.get_zip_content_as_byte_array(doc.component, "content.xml")
                styles_xml_bytes = ZIP.get_zip_content_as_byte_array(doc.component, "styles.xml")
        """
        file = str(Path(fnm))
        ox = ZipFileAccessComp.from_lo()
        storage_factory = StorageFactoryComp.from_lo()
        storage = storage_factory.create_instance()
        # next line fails.
        # stream = storage.openStorageElement("ms777", ElementModes.READWRITE)
        stream = storage.openStreamElement("zipStream", ElementModes.READWRITE)
        storable = cast("XStorable", doc)  # type: ignore
        props = mProps.Props.make_props(OutputStream=stream)
        storable.storeToURL("private:stream", props)

        ox.component.initialize((stream,))  # type: ignore
        content = ox.get_by_name(file)
        length = content.available()
        success, results = content.read_bytes(length)
        if success != length:
            raise Exception("Failed to read bytes")
        return results
