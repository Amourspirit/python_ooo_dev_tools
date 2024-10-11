from __future__ import annotations
from typing import TYPE_CHECKING, Sequence
import importlib.abc
import importlib.util
import os
import sys
from contextlib import contextmanager
from pathlib import Path
from urllib.parse import urlparse

if TYPE_CHECKING:
    import types
    from importlib.machinery import ModuleSpec
    from ooodev.utils.type_var import PathOrStr


class ImporterFile(importlib.abc.MetaPathFinder, importlib.abc.Loader):
    def __init__(self, module_path: str):
        self.module_path = module_path

    def find_spec(
        self, fullname: str, path: Sequence[str] | None, target: types.ModuleType | None = None
    ) -> ModuleSpec | None:
        # Build the path to the module file
        module_name = fullname.rsplit(".", 1)[-1]
        filename = os.path.join(self.module_path, f"{module_name}.py")
        if os.path.isfile(filename):
            return importlib.util.spec_from_file_location(fullname, filename, loader=self)
        return None

    def create_module(self, spec: ModuleSpec):
        # Use default module creation
        return None

    def exec_module(self, module: types.ModuleType):
        # Execute the module's code
        module_name = module.__name__.rsplit(".", 1)[-1]
        filename = os.path.join(self.module_path, f"{module_name}.py")
        with open(filename, "r") as file:
            source = file.read()
        code = compile(source, filename, "exec")
        exec(code, module.__dict__)

    def uri_to_path(self, file_uri: str) -> str:
        """
        Converts a file URI to a regular file path.

        Args:
            file_uri (str): The file URI to convert.

        Returns:
            str: The regular file path.
        """
        parsed_uri = urlparse(file_uri)
        return str(Path(parsed_uri.path))


@contextmanager
def importer_file(module_path: PathOrStr):
    """
    Context manager that adds ImporterUserScript to sys.meta_path
    and removes it after the context is exited.

    Args:
        module_path (PathOrStr): The path to the module to import.
    """
    importer = ImporterFile(module_path=str(Path(module_path)))
    sys.meta_path.insert(0, importer)
    try:
        yield
    finally:
        sys.meta_path.remove(importer)
