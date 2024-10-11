from __future__ import annotations
from typing import TYPE_CHECKING, Sequence
import sys
from contextlib import contextmanager
import importlib.util
import uno

# import pythonscript  # type: ignore
from ooodev.uno_helper.importer.importer_file import ImporterFile
from ooodev.loader.lo import Lo
from ooodev.io.sfa.sfa import Sfa

try:
    from typing import override  # type: ignore  # Python 3.12+
except ImportError:
    from typing_extensions import override  # For Python versions below 3.12

if TYPE_CHECKING:
    import types
    from importlib.machinery import ModuleSpec


class ImporterDocScript(ImporterFile):
    @override
    def __init__(self):
        self._sp = None
        self._sfa = Sfa()
        pv = self._get_script_provider()
        super().__init__(pv.dirBrowseNode.rootUrl)

    def _get_script_provider(self):
        """Get the user script provider."""
        if self._sp is None:
            doc = Lo.current_doc
            self._sp = doc.python_script.document_script_provider
        return self._sp

    @override
    def find_spec(
        self, fullname: str, path: Sequence[str] | None, target: types.ModuleType | None = None
    ) -> ModuleSpec | None:
        module_name = fullname.rsplit(".", 1)[-1]
        filename = f"{self.module_path}/{module_name}.py"
        if self._sfa.exists(filename):
            return importlib.util.spec_from_file_location(fullname, filename, loader=self)
        return None

    @override
    def exec_module(self, module: types.ModuleType):
        # Execute the module's code
        module_name = module.__name__.rsplit(".", 1)[-1]
        filename = f"{self.module_path}/{module_name}.py"
        source = self._sfa.read_text_file(filename)
        code = compile(source, filename, "exec")
        exec(code, module.__dict__)


@contextmanager
def importer_doc_script():
    """
    Context manager that adds ImporterUserScript to sys.meta_path
    and removes it after the context is exited.
    """
    importer = ImporterDocScript()
    sys.meta_path.insert(0, importer)
    try:
        yield
    finally:
        sys.meta_path.remove(importer)
