from __future__ import annotations
from typing import Any
from ooodev.adapter.script.document_script_library_container_comp import DocumentScriptLibraryContainerComp
from ooodev.adapter.script.document_dialog_library_container_comp import DocumentDialogLibraryContainerComp


class LibrariesPartial:
    """Partial Class used for adding access to libraries"""

    def __init__(self, component: Any) -> None:
        """
        Constructor.

        Args:
            component (Any): Any Uno Component that supports ``XStorageBasedLibraryContainer`` interface.
        """
        self.__component = component

    @property
    def basic_libraries(self) -> DocumentScriptLibraryContainerComp:
        """
        Gets the basic libraries container.

        Returns:
            DocumentScriptLibraryContainerComp: The instance.
        """
        return DocumentScriptLibraryContainerComp(self.__component.BasicLibraries)

    @property
    def dialog_libraries(self) -> DocumentDialogLibraryContainerComp:
        """
        Gets the dialog libraries container.

        Returns:
            DocumentDialogLibraryContainerComp: The instance.
        """
        return DocumentDialogLibraryContainerComp(self.__component.DialogLibraries)
