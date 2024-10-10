from __future__ import annotations
import sys
from contextlib import contextmanager
import uno

# import pythonscript  # type: ignore
from ooodev.uno_helper.py_script import python_script
from ooodev.uno_helper.importer.importer_file import ImporterFile
from ooodev.loader.lo import Lo


class ImporterUserScript(ImporterFile):
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

    def find_spec(self, fullname, path, target=None):
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
