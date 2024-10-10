from __future__ import annotations
import importlib.abc
import importlib.util
import os
import sys
from contextlib import contextmanager
from pathlib import Path
from urllib.parse import urlparse


class ImporterFile(importlib.abc.MetaPathFinder, importlib.abc.Loader):
    def __init__(self, module_path):
        self.module_path = module_path

    def find_spec(self, fullname, path, target=None):
        # Build the path to the module file
        module_name = fullname.rsplit(".", 1)[-1]
        filename = os.path.join(self.module_path, f"{module_name}.py")
        if os.path.isfile(filename):
            return importlib.util.spec_from_file_location(fullname, filename, loader=self)
        return None

    def create_module(self, spec):
        # Use default module creation
        return None

    def exec_module(self, module):
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
def importer_file_context(module_path: str):
    """
    Context manager that adds ImporterUserScript to sys.meta_path
    and removes it after the context is exited.
    """
    importer = ImporterFile(module_path=module_path)
    sys.meta_path.insert(0, importer)
    try:
        yield
    finally:
        sys.meta_path.remove(importer)
