from math import e
from multiprocessing import Value
import sys
import inspect
import types
import typing
import ast
from contextlib import contextmanager
from ast import parse, NodeVisitor, ImportFrom
from importlib import util as import_util, import_module, reload
from importlib.machinery import ModuleSpec
from os import path
from pkgutil import iter_modules
from typing import Any, List, Iterator

import pytest

# https://medium.com/brexeng/avoiding-circular-imports-in-python-7c35ec8145ed
# https://github.com/amenck/circular_import_refactor
# https://github.com/mitmproxy/pdoc/blob/main/pdoc/doc_ast.py


IF_IGNORE_ATTRIBUTES = {"TYPE_CHECKING", "DOCS_BUILDING"}
MODULES_EXCLUDE = {"ooodev.dialog.tk_input", "ooodev.utils.color"}
MODULES_PREFIX_EXCLUDE = ("ooodev.adapter._helper",)


@contextmanager
def patch(object_, attribute_name, value):
    old_value = getattr(object_, attribute_name)
    try:
        setattr(object_, attribute_name, value)
        yield
    finally:
        setattr(object_, attribute_name, old_value)


def detect_type_checking_mode_modules_names(module_name):
    # https://stackoverflow.com/questions/54764937/detecting-modules-that-get-imported-on-condition
    # this method loads the module as default (Type checking false). Gets then module vars,
    # then turns on type checking and reloads the module to get the type checked vars.
    module = import_module(module_name)
    default_module_names = set(vars(module))

    with patch(typing, "TYPE_CHECKING", True):
        # reloading since ``importlib.import_module``
        # will return previously cached entry
        reload(module)
    type_checked_module_namespace = dict(vars(module))

    # resetting to "default" mode
    reload(module)

    return {
        name
        for name, content in type_checked_module_namespace.items()
        if name not in default_module_names and inspect.ismodule(content)
    }


def _is_test_module(module_name: str) -> bool:
    components = module_name.split(".")

    return len(components) >= 2 and components[1] == "tests"


def _is_package(module_spec: ModuleSpec) -> bool:
    return module_spec.origin is not None and module_spec.origin.endswith("__init__.py")


def _recurse_modules(module_name: str, ignore_tests: bool, packages_only: bool) -> Iterator[str]:
    if ignore_tests and _is_test_module(module_name):
        return

    module_spec = import_util.find_spec(module_name)

    if module_spec is not None and module_spec.origin is not None:
        if not (packages_only and not _is_package(module_spec)):
            yield module_name

        for child in iter_modules([path.dirname(module_spec.origin)]):
            if child.ispkg:
                yield from _recurse_modules(
                    f"{module_name}.{child.name}",
                    ignore_tests=ignore_tests,
                    packages_only=packages_only,
                )
            elif not packages_only:
                yield f"{module_name}.{child.name}"


class _ImportFromSourceChecker(NodeVisitor):
    def __init__(self, module: str):
        module_spec = import_util.find_spec(module)
        is_pkg = (
            module_spec is not None and module_spec.origin is not None and module_spec.origin.endswith("__init__.py")
        )

        self._module = module if is_pkg else ".".join(module.split(".")[:-1])
        self._top_level_module = self._module.split(".")[0]
        if module_spec is not None and hasattr(module_spec, "origin"):
            self._module_file = module_spec.origin  # type: ignore
        else:
            self._module_file = ""

    def _test_if_module_import(self, module_to_import: str) -> bool:
        try:
            module = import_module(module_to_import)
        except ImportError:
            return False
        if module_to_import != module.__name__:
            raise ValueError(
                f"Imported Module {module_to_import} does not match actual module name: {module.__name__}"
            )
        return True

    def visit_ImportFrom(self, node: ImportFrom) -> Any:
        # Check that there are no relative imports that attempt to read from a parent module. We've found that there
        # generally is no good reason to have such imports.
        if node.level >= 2:
            raise ValueError(
                f"{self._module_file} "
                f"Import in {self._module} attempts to import from parent module using relative import. Please "
                f"switch to absolute import instead."
            )

        # Figure out which module to import in the case where this is a...
        if node.level == 0:
            # (1) absolute import where a submodule is specified
            assert node.module is not None
            module_to_import: str = node.module
        elif node.module is None:
            # (2) relative import where no module is specified (ie: "from . import foo")
            module_to_import = self._module
        else:
            # (3) relative import where a submodule is specified (ie: "from .bar import foo")
            module_to_import = f"{self._module}.{node.module}"

        # We're only looking at imports of objects defined inside this top-level package
        if not module_to_import.startswith(self._top_level_module):
            return

        if module_to_import.startswith(MODULES_PREFIX_EXCLUDE):
            # if a module has a path part that starts with an underscore it will not work for this test.
            return
        if module_to_import in MODULES_EXCLUDE:
            return
        # Actually import the module and iterate through all the objects potentially exported by it.
        module = import_module(module_to_import)
        for alias in node.names:
            # if getting an error here it may be because then module has been imported using a
            # module alias.
            # for instance the ooodev.format.inner.direct.write.table.props.table_properties contains a bunch of classes.
            # the module is imported as tp. So the alias name is tp.
            # from ooodev.format.inner.direct.write.table.props import table_properties as tp
            # this is a problem as the module imported would ooodev.format.inner.direct.write.table.props which is a module package (__init__.py)
            # therefore hasattr(module, alias.name) would return False.
            # then solution is to not import modules aliases an instead import the module classes directly. The classes can be imported as aliases.
            # from ooodev.format.inner.direct.write.table.props.table_properties import (
            #     TblAbsUnit,
            #     TblRelUnit,
            #     TableAlignKind,
            #     TableProperties,
            # )
            # see: ooodev.format.inner.partial.write.table.write_table_properties_partial.WriteTablePropertiesPartial
            if not hasattr(module, alias.name):
                try:
                    if self._test_if_module_import(f"{module_to_import}.{alias.name}"):
                        continue
                except Exception as e:
                    raise ValueError(
                        f"Test alias.name of {alias.name} as module has failed: {self._module_file}. {self._module_file}"
                    ) from e

                raise ValueError(
                    f"NO alias name attr: {self._module_file}. {self._module_file} "
                    f"Imported {alias.name} from {module_to_import}."
                )
            attr = getattr(module, alias.name)

            # For some objects (pretty much everything except for classes and functions), we are not able to figure
            # out which module they were defined in... in that case there's not much we can do here, since we cannot
            # easily figure out where we *should* be importing this from in the first place.
            if attr is Any:
                continue
            if isinstance(attr, type) or callable(attr):
                attribute_module = attr.__module__
            else:
                continue

            # Figure out where we should be importing this class from, and assert that the *actual* import we found
            # matches the place we *should* import from.
            should_import_from = self._get_module_should_import(module_to_import=attribute_module)
            assert module_to_import == should_import_from, (
                f"{self._module_file} "
                f"Imported {alias.name} from {module_to_import}, which is not the public module where this object "
                f"is defined. Please import from {should_import_from} instead."
            )

    def _get_module_should_import(self, module_to_import: str) -> str:
        """
        This function figures out the correct import path for "module_to_import" from the "self._module" module in
        this instance. The trivial solution here would be to always just return "module_to_import", but we want
        to actually take into account the fact that some submodules can be "private" (ie: start with an "_"), in
        which case we should only import from them if self._module is internal to that private module.
        """
        module_components = module_to_import.split(".")
        result: List[str] = []

        for component in module_components:
            if component.startswith("_") and not self._module.startswith(".".join(result)):
                break
            result.append(component)

        return ".".join(result)


def _apply_visitor(module: str, visitor: NodeVisitor) -> None:
    module_spec = import_util.find_spec(module)
    assert module_spec is not None
    assert module_spec.origin is not None

    with open(module_spec.origin, "r") as source_file:
        ast_mod = parse(source=source_file.read(), filename=module_spec.origin)

    # remove any imports that are guarded by type checking.
    # Without this a lot of issues will be reported that are not relevant.
    _parse_ast_remove_type_checking_from_imports(ast_mod)
    visitor.visit(ast_mod)


def _parse_ast_remove_type_checking_from_imports(mod: ast.Module) -> None:
    """
    Walks the abstract syntax tree for `mod` and removes all from imports statements guarded by TYPE_CHECKING blocks.
    """

    def parse_typing_block(if_node: ast.If) -> None:
        remove_indexes = [i for i, node in enumerate(if_node.body) if isinstance(node, ast.ImportFrom)]
        remove_indexes.reverse()
        for i in remove_indexes:
            if_node.body.pop(i)

        # remove_indexes = [i for i, node in enumerate(if_node.orelse) if isinstance(node, ast.ImportFrom)]
        # remove_indexes.reverse()
        # for i in remove_indexes:
        #     if_node.orelse.pop(i)

    for node in mod.body:
        if isinstance(node, ast.If) and isinstance(node.test, ast.Name) and node.test.id == "TYPE_CHECKING":
            parse_typing_block(node)
        if (
            isinstance(node, ast.If)
            and isinstance(node.test, ast.Attribute)
            and isinstance(node.test.value, ast.Name)
            # some folks do "import typing as t", the accuracy with just TYPE_CHECKING is good enough.
            # and node.test.value.id == "typing"
            and node.test.attr in IF_IGNORE_ATTRIBUTES
        ):
            parse_typing_block(node)


def _test_imports_from_source(module: str) -> None:
    if not module:
        return
    _apply_visitor(module=module, visitor=_ImportFromSourceChecker(module))


def add_module_organization_tests(module_name: str) -> None:
    """
    This function dynamically generates a set of python tests which can be used to ensure that all modules follow
    the convention "classes and functions should always be imported from the module they are defined in (or the
    closest public module to that)".

    Let's say that you have a package "foo", and want to use this function. In that case, go into your test module
    (probably "foo.tests") and create a test file that imports and calls `add_module_organization_tests(__name__)`.
    Once this is defined, you can use pytest to run your tests, and note that a unique test has been generated for
    each file in your project. The tests will scan each file and look for cases of imports that do not follow the
    convention above. If the test finds any violations, it will error out with a message similar to:

    AssertionError: Imported Child from objects, which is not the public module where this object is defined. Please
        import from objects.child instead.
    """
    module_root = module_name.split(".")[0]
    lst = list(_recurse_modules(module_root, ignore_tests=True, packages_only=False))
    setattr(
        sys.modules[module_name],
        "test_imports_from_source",
        pytest.mark.parametrize(
            "module",
            lst,
        )(_test_imports_from_source),
    )


def get_modules(module_name: str) -> List[str]:
    module_root = module_name.split(".")[0]
    return list(_recurse_modules(module_root, ignore_tests=True, packages_only=False))
