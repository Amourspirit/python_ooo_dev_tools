from __future__ import annotations
from typing import TYPE_CHECKING, Sequence
import sys
from contextlib import contextmanager
import uno

from ooodev.uno_helper.importer.importer_file import ImporterFile
from ooodev.loader.lo import Lo

try:
    from typing import override  # type: ignore  # Python 3.12+
except ImportError:
    from typing_extensions import override  # For Python versions below 3.12

if TYPE_CHECKING:
    import types
    from importlib.machinery import ModuleSpec


class ImporterSharedExtScript(ImporterFile):
    @override
    def __init__(self, ext_name: str):
        """
        Initializes the ImporterSharedExtScript instance.
        Args:
            ext_name (str): The name of the extension module to be imported.
        Raises:
            ImportError: If the module specified by `ext_name` is not found.
        """

        self._sp = None
        self._ext_name = ext_name
        self._location = None
        self.__module_path = ""
        found = self._find_module()
        if not found:
            raise ImportError(f"Module {ext_name} not found")
        super().__init__(self.__module_path)

    def _find_module(self):

        sp = Lo.current_doc.python_script.shared_script_provider.uno_packages_sp
        if self._search_node(sp, self._ext_name, ".oxt"):
            return True
        return False

    def _search_node(self, node, name, ext=""):
        for child in node.getChildNodes():
            if child.name == uno.systemPathToFileUrl(name + ext):
                self.__module_path = self.uri_to_path(child.rootUrl)
                return True
        return False

    @override
    def find_spec(
        self, fullname: str, path: Sequence[str] | None, target: types.ModuleType | None = None
    ) -> ModuleSpec | None:
        if fullname.startswith("com."):
            return None
        return super().find_spec(fullname, path, target)


@contextmanager
def importer_shared_ext_script_context(ext_name: str):
    """
    Context manager that adds ImporterSharedExtScript to sys.meta_path
    and removes it after the context is exited.
    """
    importer = ImporterSharedExtScript(ext_name)
    sys.meta_path.insert(0, importer)
    try:
        yield
    finally:
        sys.meta_path.remove(importer)
