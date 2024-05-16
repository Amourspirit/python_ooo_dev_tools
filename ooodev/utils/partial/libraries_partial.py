from __future__ import annotations
from typing import Any, TYPE_CHECKING
from ooodev.mock import mock_g
from ooodev.adapter.script.document_script_library_container_comp import DocumentScriptLibraryContainerComp
from ooodev.adapter.script.document_dialog_library_container_comp import DocumentDialogLibraryContainerComp
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial

if TYPE_CHECKING:
    from ooodev.macro.script.python_script import PythonScript


class LibrariesPartial:
    """Partial Class used for adding access to libraries"""

    def __init__(self, component: Any) -> None:
        """
        Constructor.

        Args:
            component (Any): Any Uno Component that supports ``XStorageBasedLibraryContainer`` interface.
        """
        self.__component = component
        self.__python_script = None
        self.__scripts = None
        self.__dialogs = None

    @property
    def basic_libraries(self) -> DocumentScriptLibraryContainerComp:
        """
        Gets the basic libraries container.

        Returns:
            DocumentScriptLibraryContainerComp: The instance.
        """
        if self.__scripts is None:
            self.__scripts = DocumentScriptLibraryContainerComp(self.__component.BasicLibraries)
        return self.__scripts

    @property
    def dialog_libraries(self) -> DocumentDialogLibraryContainerComp:
        """
        Gets the dialog libraries container.

        Returns:
            DocumentDialogLibraryContainerComp: The instance.
        """
        if self.__dialogs is None:
            self.__dialogs = DocumentDialogLibraryContainerComp(self.__component.DialogLibraries)
        return self.__dialogs

    @property
    def python_script(self) -> PythonScript:
        """
        Gets the python script instance.

        Returns:
            PythonScript: The instance.
        """
        if self.__python_script is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.macro.script.python_script import PythonScript

            lo_inst = None
            if isinstance(self, LoInstPropsPartial):
                lo_inst = self.lo_inst
            self.__python_script = PythonScript(doc=self.__component, lo_inst=lo_inst)  # type: ignore
        return self.__python_script


if mock_g.FULL_IMPORT:
    from ooodev.macro.script.python_script import PythonScript
