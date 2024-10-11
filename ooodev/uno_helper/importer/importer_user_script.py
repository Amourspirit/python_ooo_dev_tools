from __future__ import annotations
from typing import TYPE_CHECKING, Sequence
import sys
from contextlib import contextmanager
import uno

# import pythonscript  # type: ignore
from ooodev.uno_helper.py_script import python_script
from ooodev.uno_helper.importer.importer_file import ImporterFile
from ooodev.loader.lo import Lo

try:
    from typing import override  # type: ignore  # Python 3.12+
except ImportError:
    from typing_extensions import override  # For Python versions below 3.12

if TYPE_CHECKING:
    import types
    from importlib.machinery import ModuleSpec


class ImporterUserScript(ImporterFile):
    @override
    def __init__(self):
        self._sp = None
        pv = self._get_script_provider()
        super().__init__(self.uri_to_path(pv.dirBrowseNode.rootUrl))

    def _get_script_provider(self):
        """Get the user script provider."""
        if self._sp is None:
            ctx = Lo.get_context()  #  uno.getComponentContext()
            PythonScriptProvider = python_script.PythonScriptProvider  # type: ignore
            self._sp = PythonScriptProvider(ctx, "user")
            try:
                self._sp.uno_packages_sp = PythonScriptProvider(ctx, "user:uno_packages")
            except Exception:
                self._sp.uno_packages_sp = None
        return self._sp

    @override
    def find_spec(
        self, fullname: str, path: Sequence[str] | None, target: types.ModuleType | None = None
    ) -> ModuleSpec | None:
        if fullname.startswith("com."):
            return None
        return super().find_spec(fullname, path, target)


@contextmanager
def importer_user_script_context():
    """
    Context manager that adds ImporterUserScript to sys.meta_path
    and removes it after the context is exited.
    """
    importer = ImporterUserScript()
    sys.meta_path.insert(0, importer)
    try:
        yield
    finally:
        sys.meta_path.remove(importer)
