from __future__ import annotations
import sys
import os
import inspect
from typing import Any, Tuple, List
from types import ModuleType
import uno


# consider * imports such as from ooodev.utils.builder.dynamic_importer import *

# uno import notes.
# uno imports such as form com.sun.star.awt import XComboBox are not added to the python sys.modules.
# This is because uno hooks the __import__ function and imports the classes directly.
# For this reason is does not seem possible to unload the module from sys.modules.


class DynamicImporter:
    def __init__(self, module_name: str, class_name: str = "", alias: str = ""):
        self._module_name = module_name
        self._class_name = class_name
        self._alias = alias

    def load_module(self) -> None | Tuple[ModuleType, Any]:
        result, mod, clz = self._load_import()
        if not result:
            return None

        return mod, clz

    def is_module_existing_import(self) -> bool:
        if self.is_uno_import:
            return False
        return self._module_name in sys.modules

    def unload_module(self) -> bool:
        if self.is_uno_import:
            return False
        if self.is_module_existing_import():
            del sys.modules[self._module_name]
            return True
        return False

    def get_class_name(self, obj: Any) -> str:
        if inspect.isclass(obj) is False:
            obj = obj.__class__
        class_name = obj.__name__
        return class_name

    # Works both for python 2 and 3
    def get_class_names(self, obj: Any, include_obj: bool = True) -> List[str]:
        class_name = []
        if inspect.isclass(obj) is False:
            obj = obj.__class__
        classes = inspect.getmro(obj)
        c_name = str(obj.__name__)
        for cl in classes:
            name = cl.__name__
            if include_obj is False:
                if name == c_name:
                    continue
        return class_name

    def _load_import(self, raise_err: bool = False) -> Tuple[bool, ModuleType, Any]:
        """
        Load an import name for later comparison.
        Name is expected in format of ``scratch.uno_obj.form.component.rich_text_control.RichTextControl``

        Args:
            imports (List[str]): List of imports

        Returns:
            Tuple[bool, ModuleType, Any]: Tuple of boolean, ModuleType, and class object
        """
        try:
            im = self.full_import
            parts = im.split(sep=".")
            if self._class_name:
                cls_name = parts.pop()
            else:
                cls_name = ""
            if len(parts) > 1:
                mod_name = parts.pop()
            else:
                mod_name = ""
            pkg_name = ".".join(parts)
            mt, cl = self._dynamic_imp(pkg_name, mod_name, cls_name)
            return True, mt, cl
            # names_lst = self.get_class_names(cl, False)
        except Exception:
            if raise_err:
                raise
            return False, None, None  # type: ignore

    def _dynamic_imp(self, package: str, mod_name: str, class_name: str) -> Tuple[ModuleType, Any]:
        try:
            # __import__ needs to be used here.
            # The reason for thi is uno hook the __import__ and import uno classes.
            # if importlib is used, it will not be able to import uno classes.
            if mod_name:
                key = f"{package}.{mod_name}"
            else:
                key = package

            if class_name:
                module = __import__(key, fromlist=[class_name])
            else:
                module = __import__(key)

            if class_name:
                my_class = getattr(module, class_name)
            else:
                my_class = None
            return module, my_class
        except Exception as e:
            # print(e)
            raise e

    @property
    def has_alias(self):
        return bool(self._alias)

    @property
    def full_import(self):
        if self._class_name:
            return f"{self._module_name}.{self._class_name}"
        return self._module_name

    @property
    def is_uno_import(self):
        return self._module_name.startswith("com.sun.star.")

    @staticmethod
    def from_import(from_import: str, alias: str = "") -> DynamicImporter:
        """
        Use with building a from import statement.
        Such as ``from_import("ooodev.utils.builder.dynamic_importer.DynamicImporter")``

        Would build an import statement of ``from ooodev.utils.builder.dynamic_importer import DynamicImporter``

        Args:
            from_import (str): Import String, The last part is the class name.
            alias (str, optional): When provided the alias is passed to the instance. Defaults to "".

        Returns:
            DynamicImporter: Instance of DynamicImporter
        """
        # from ooodev.utils.builder.dynamic_importer import DynamicImporter
        module_name, class_name = from_import.rsplit(".", 1)
        return DynamicImporter(module_name, class_name, alias)

    @staticmethod
    def from_direct_import(direct_import: str, alias="") -> DynamicImporter:
        """
        Direct import of a module and class.
        Such as ``from_direct_import("math")``

        Would result in an import statement of ``import math``.

        Args:
            direct_import (str): Import String.

        Returns:
            DynamicImporter: Instance of DynamicImporter
        """
        return DynamicImporter(direct_import, alias=alias)
