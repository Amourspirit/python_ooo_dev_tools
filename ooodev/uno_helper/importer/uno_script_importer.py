from __future__ import annotations
from typing import Any
import sys
import uno
import contextlib
import os
import types
from pathlib import Path

import pythonscript  # type: ignore
from ooodev.loader.lo import Lo
from ooodev.io.log.named_logger import NamedLogger


class UnoScriptImporter:
    """Allow the user to import, from modules located in containers"""

    def __init__(self):
        self._ctx = Lo.current_lo.get_context()  # uno.getComponentContext()
        self._providers = self._load_providers()
        self._nodes = {}
        self._log = NamedLogger(self.__class__.__name__)
        self._log.debug("__init__")

    def _load_providers(self) -> dict:
        p = {}
        p["user"] = Lo.current_doc.python_script.user_script_provider, ""
        p["user:uno_packages"] = Lo.current_doc.python_script.user_script_provider.uno_packages_sp, ".oxt"
        p["share"] = Lo.current_doc.python_script.shared_script_provider, ""
        p["share:uno_packages"] = Lo.current_doc.python_script.shared_script_provider.uno_packages_sp, ".oxt"

        with contextlib.suppress(AttributeError, NameError):
            doc = Lo.current_lo.current_doc
            if doc is None:
                self._log.warning("_load_providers() No current document.")
                return p
            if doc.component.ScriptContainer:
                p["document"] = (Lo.current_doc.python_script.document_script_provider, "")
        return p

    def _find_module(self, name: str, path: str):
        if self._log.is_debug:
            self._log.debug("_find_module()")
            if path:
                self._log.debug(f"{name=}, {path=}")
            else:
                self._log.debug(f"{name=}:")
        if path:
            node = path[0]
            if node in self._nodes:
                return self._search_node(self._nodes[node], name)
            else:
                return False
        else:
            for prov in self._providers:
                sp, ext = self._providers[prov]
                if self._search_node(sp, name, ext):
                    self.location = prov
                    self._log.debug(f"Found in {prov}.")
                    return True
                self._log.debug(f"Not found in {prov}.")
        return False

    def _search_node(self, node, name, ext=""):
        self._log.debug("searching...")
        for child in node.getChildNodes():
            if child.name == uno.systemPathToFileUrl(name + ext):
                self._nodes[self.fullname] = child
                return True
        return False

    def find_module(self, fullname: str, path: Any = None) -> UnoScriptImporter | None:
        """
        Attempts to find a module by its full name.

        Args:
            fullname (str): The full name of the module to find.
            node_path (str, optional): The node path to search for the module. Defaults to an empty string.

        Returns:
            UnoScriptImporter: Returns the importer instance if the module is found, otherwise None.
        """

        self._log.debug("UnoScriptImporter.find_module")
        self._log.debug(f"fullname: {fullname}")
        if not fullname:
            return None
        if fullname.startswith("com."):
            return None
        self.fullname = fullname
        name = fullname.rsplit(".", 1)[-1]
        return self if self._find_module(name, path) else None

    # def _module_from_node(self, node: Any):
    #     import imp

    #     mod = imp.new_module(node.name)
    #     for child in node.getChildNodes():
    #         try:
    #             if isinstance(child, pythonscript.FileBrowseNode):
    #                 setattr(mod, child.name, node.provCtx.getModuleByUrl(child.uri))
    #             else:
    #                 setattr(mod, child.name, self._module_from_node(child))
    #         except ImportError:
    #             self._log.exception(f"Unexpected error while loading module <{child.name}>")
    #     return mod

    def _module_from_node(self, node: Any):
        mod = types.ModuleType(node.name)
        for child in node.getChildNodes():
            try:
                if isinstance(child, pythonscript.FileBrowseNode):
                    setattr(mod, child.name, node.provCtx.getModuleByUrl(child.uri))
                else:
                    setattr(mod, child.name, self._module_from_node(child))
            except ImportError:
                self._log.exception(f"Unexpected error while loading module <{child.name}>")
        return mod

    def load_module(self, fullname: str) -> Any:
        """
        Loads a module given its full name.
        This method attempts to load a module by its full name.
        It first logs the attempt and then retrieves the corresponding node from the internal nodes dictionary.
        Depending on the type of node, it either constructs the module from the node or retrieves it using the provided URL.
        The method also sets various attributes on the module, such as ``__file__``, ``__path__``, ``__package__``, ``__name__``, and ``__loader__``.
        Finally, it registers the module in the ``sys.modules`` dictionary and returns it.

        Args:
            fullname (str): The full name of the module to load.

        Returns:
            Any: The loaded module.
        """

        self._log.debug("load_module()")
        self._log.debug(f"fullname: {fullname}")
        node = self._nodes[fullname]
        if isinstance(node, pythonscript.DirBrowseNode):
            mod = self._module_from_node(node)
            mod.__file__ = f"<{self.location}>"
            mod.__path__ = [fullname]
            mod.__package__ = fullname
        else:
            mod = node.provCtx.getModuleByUrl(node.uri)
        with contextlib.suppress(IndexError):
            parent = fullname.rsplit(".", 1)[-2]
            mod.__package__ = parent
        mod.__name__ = fullname
        mod.__loader__ = self
        sys.modules[fullname] = mod
        return mod


class UnoScriptImporter2:
    """Allow the user to import, from modules located in containers"""

    def __init__(self):
        self._ctx = Lo.current_lo.get_context()  # uno.getComponentContext()
        self._providers = self._load_providers()
        self._nodes = {}
        self._log = NamedLogger(self.__class__.__name__)
        self._log.debug("__init__")

    def _load_providers(self) -> dict:
        p = {}
        p["user"] = Lo.current_doc.python_script.user_script_provider, ""
        p["user:uno_packages"] = Lo.current_doc.python_script.user_script_provider, "ext"
        p["share"] = Lo.current_doc.python_script.shared_script_provider, ""
        p["share:uno_packages"] = Lo.current_doc.python_script.shared_script_provider, "ext"

        with contextlib.suppress(AttributeError, NameError):
            doc = Lo.current_lo.current_doc
            if doc is None:
                self._log.warning("_load_providers() No current document.")
                return p
            if doc.component.ScriptContainer:
                p["document"] = (Lo.current_doc.python_script.document_script_provider, "")
        return p

    def _load_providers2(self) -> dict:
        p = {}
        locations = {"user", "user:uno_packages", "share", "share:uno_packages"}
        for location in locations:
            ext = ""
            if location.endswith("uno_packages"):
                ext = ".oxt"
            p[location] = (pythonscript.PythonScriptProvider(self._ctx, location), ext)
        with contextlib.suppress(AttributeError, NameError):
            doc = Lo.current_lo.current_doc
            if doc is None:
                self._log.warning("_load_providers() No current document.")
                return p
            if doc.component.ScriptContainer:
                p["document"] = (pythonscript.PythonScriptProvider(self._ctx, doc.component), "")
        return p

    def _find_module(self, name: str, node_path: str):
        if self._log.is_debug:
            self._log.debug("_find_module()")
            if node_path:
                self._log.debug(f"{name=}, {node_path=}")
            else:
                self._log.debug(f"{name=}:")
        if node_path:
            node = node_path[0]
            if node in self._nodes:
                return self._search_node(self._nodes[node], name)
            else:
                return False
        else:
            for prov in self._providers:
                sp, ext = self._providers[prov]
                if self._search_node(sp, name, ext):
                    self.location = prov
                    self._log.debug(f"Found in {prov}.")
                    return True
                self._log.debug(f"Not found in {prov}.")
        return False

    def _search_node(self, node, name, ext=""):
        self._log.debug("searching...")
        for child in node.getChildNodes():
            if child.name == uno.systemPathToFileUrl(name + ext):
                self._nodes[self.fullname] = child
                return True
        return False

    def find_module(self, fullname: str, node_path: str = "") -> UnoScriptImporter | None:
        """
        Attempts to find a module by its full name.

        Args:
            fullname (str): The full name of the module to find.
            node_path (str, optional): The node path to search for the module. Defaults to an empty string.

        Returns:
            UnoScriptImporter: Returns the importer instance if the module is found, otherwise None.
        """

        self._log.debug("UnoScriptImporter.find_module")
        self._log.debug(f"fullname: {fullname}")
        if fullname in {"com"}:
            return None
        self.fullname = fullname
        name = fullname.rsplit(".", 1)[-1]
        return self if self._find_module(name, node_path) else None

    def _module_from_node(self, node: Any):
        import imp

        mod = imp.new_module(node.name)
        for child in node.getChildNodes():
            try:
                if isinstance(child, pythonscript.FileBrowseNode):
                    setattr(mod, child.name, node.provCtx.getModuleByUrl(child.uri))
                else:
                    setattr(mod, child.name, self._module_from_node(child))
            except ImportError:
                self._log.exception(f"Unexpected error while loading module <{child.name}>")
        return mod

    def load_module(self, fullname: str) -> Any:
        """
        Loads a module given its full name.
        This method attempts to load a module by its full name.
        It first logs the attempt and then retrieves the corresponding node from the internal nodes dictionary.
        Depending on the type of node, it either constructs the module from the node or retrieves it using the provided URL.
        The method also sets various attributes on the module, such as ``__file__``, ``__path__``, ``__package__``, ``__name__``, and ``__loader__``.
        Finally, it registers the module in the ``sys.modules`` dictionary and returns it.

        Args:
            fullname (str): The full name of the module to load.

        Returns:
            Any: The loaded module.
        """

        self._log.debug("load_module()")
        self._log.debug(f"fullname: {fullname}")
        node = self._nodes[fullname]
        if isinstance(node, pythonscript.DirBrowseNode):
            mod = self._module_from_node(node)
            mod.__file__ = f"<{self.location}>"
            mod.__path__ = [fullname]
            mod.__package__ = fullname
        else:
            mod = node.provCtx.getModuleByUrl(node.uri)
        with contextlib.suppress(IndexError):
            parent = fullname.rsplit(".", 1)[-2]
            mod.__package__ = parent
        mod.__name__ = fullname
        mod.__loader__ = self
        sys.modules[fullname] = mod
        return mod
