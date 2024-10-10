from __future__ import annotations

import pythonscript  # type: ignore
from ooodev.loader.lo import Lo

from ooodev.uno_helper.importer.importer_file import ImporterFile


class UserScriptImporter(ImporterFile):
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

    def find_spec(self, fullname, path, target=None):
        if fullname.startswith("com."):
            return None
        return super().find_spec(fullname, path, target)
