from __future__ import annotations
from typing import TYPE_CHECKING, Sequence
import pythonscript  # type: ignore
from ooodev.loader.lo import Lo

from ooodev.uno_helper.importer.importer_file import ImporterFile

try:
    from typing import override  # type: ignore  # Python 3.12+
except ImportError:
    from typing_extensions import override  # For Python versions below 3.12

if TYPE_CHECKING:
    import types
    from importlib.machinery import ModuleSpec


class UserScriptImporter(ImporterFile):
    @override
    def __init__(self):
        self._sp = None
        pv = self._get_script_provider()
        super().__init__(self.uri_to_path(pv.dirBrowseNode.rootUrl))

    def _get_script_provider(self):
        """Get the share script provider."""
        if self._sp is None:
            ctx = Lo.current_lo.get_context()
            PythonScriptProvider = pythonscript.PythonScriptProvider  # type: ignore
            self._sp = PythonScriptProvider(ctx, "share")
            try:
                self._sp.uno_packages_sp = PythonScriptProvider(ctx, "share:uno_packages")
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
