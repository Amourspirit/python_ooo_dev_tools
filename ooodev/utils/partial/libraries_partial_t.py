from __future__ import annotations
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ooodev.adapter.script.document_script_library_container_comp import DocumentScriptLibraryContainerComp
    from ooodev.adapter.script.document_dialog_library_container_comp import DocumentDialogLibraryContainerComp
    from ooodev.macro.script.python_script import PythonScript

    from typing_extensions import Protocol
else:
    Protocol = object


class LibrariesPartialT(Protocol):

    @property
    def basic_libraries(self) -> DocumentScriptLibraryContainerComp:
        """
        Gets the basic libraries container.

        Returns:
            DocumentScriptLibraryContainerComp: The instance.
        """
        ...

    @property
    def dialog_libraries(self) -> DocumentDialogLibraryContainerComp:
        """
        Gets the dialog libraries container.

        Returns:
            DocumentDialogLibraryContainerComp: The instance.
        """
        ...

    @property
    def python_script(self) -> PythonScript:
        """
        Gets the python script instance.

        Returns:
            PythonScript: The instance.
        """
        ...
