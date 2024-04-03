from __future__ import annotations
from typing import Any, Dict, List, Type, Set, TYPE_CHECKING
import copy
import importlib

from ooodev.utils.builder.build_import_arg import BuildImportArg
from ooodev.loader import lo as mLo
from ooodev.utils import info as mInfo
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.builder.init_kind import InitKind
from ooodev.utils.builder.check_kind import CheckKind
from ooodev.adapter.component_base import ComponentBase

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst


class DefaultBuilder(ComponentBase, LoInstPropsPartial):
    def __init__(self, component: Any, lo_inst: LoInst | None = None):
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst)
        ComponentBase.__init__(self, component)

        self._component = component
        # _bases_partial could be classes such as property classes
        self._bases_partial: Dict[Type, BuildImportArg] = {}
        # _bases_interfaces could be interfaces such as ComponentPartial
        self._bases_interfaces: Dict[Type, BuildImportArg] = {}
        # _build_args is a dictionary of BuildImportArg, using dictionary to avoid duplicates and to keep order
        self._build_args: Dict[BuildImportArg, Any] = {}
        self._omit: Set[str] = set()

    def _get_import(self, name: str) -> Any:
        return importlib.import_module(name)

    def _get_mod_class_names(self, name: str) -> tuple[str, str]:
        return tuple(name.rsplit(".", 1))  # type: ignore

    def _get_class(self, arg: BuildImportArg) -> Any:
        if not self._passes_check(arg):
            raise ValueError(f"Component does not support {arg.uno_name}")
        mod_name, class_name = self._get_mod_class_names(arg.ooodev_name)
        return getattr(self._get_import(mod_name), class_name)

    def _has_interface(self, name: str) -> bool:
        return mLo.Lo.is_uno_interfaces(self._component, name)

    def _supports_service(self, name: str) -> bool:
        return mInfo.Info.support_service(self._component, name)

    def _get_optional_class(self, arg: BuildImportArg) -> Any:
        if not arg.uno_name:
            return None
        if self._passes_check(arg):
            return self._get_class(arg)
        return None

    def _passes_check(self, arg: BuildImportArg) -> bool:
        if not arg.uno_name:
            return True
        if arg.check_kind == CheckKind.SERVICE:
            return self._supports_service(arg.uno_name)
        elif arg.check_kind == CheckKind.INTERFACE:
            return self._has_interface(arg.uno_name)
        return True

    def _add_base(self, base: Type[Any], arg: BuildImportArg) -> None:
        if arg.ooodev_name in self._omit:
            return
        if arg.uno_name and arg.uno_name in self._omit:
            return

        if arg.import_kind == InitKind.COMPONENT:
            if self._passes_check(arg):
                self._bases_partial[base] = arg
        else:
            if self._passes_check(arg):
                self._bases_interfaces[base] = arg

    def _clear_imports(self) -> None:
        self._bases_partial.clear()
        self._bases_interfaces.clear()

    def _process_imports(self) -> None:
        self._clear_imports()
        for arg in self._build_args.keys():
            if arg.optional:
                clz = self._get_optional_class(arg)
                if clz:
                    self._add_base(clz, arg)
            else:
                clz = self._get_class(arg)
                self._add_base(clz, arg)

    def _init_class(self, instance: Any, clz: Type[Any], kind: InitKind) -> Any:
        if kind == InitKind.COMPONENT:
            clz.__init__(instance, self._component)  # type: ignore
        elif kind == InitKind.COMPONENT_INTERFACE:
            clz.__init__(instance, self._component, None)  # type: ignore
        else:
            clz.__init__(instance)

    def _create_class(self, clz: Type[Any], kind: InitKind) -> Any:
        if kind == InitKind.COMPONENT:
            inst = clz(self._component)  # type: ignore
        elif kind == InitKind.COMPONENT_INTERFACE:
            inst = clz(self._component, None)  # type: ignore
        else:
            inst = clz()
        return inst

    def _init_classes(self, instance: Any) -> None:

        for clz, kind in self._bases_partial.items():
            self._init_class(instance, clz, kind.import_kind)
        for clz, kind in self._bases_interfaces.items():
            self._init_class(instance, clz, kind.import_kind)

    def _generate_class(self, base_class: Type[Any], name: str, **kwargs) -> Type[Any]:
        bases = [base_class] + list(self._bases_interfaces.keys()) + list(self._bases_partial.keys())
        if len(bases) == 1 and not kwargs:
            return base_class
        else:
            return type(name, tuple(bases), kwargs)

    def add_build_arg(self, *args: BuildImportArg) -> None:
        for arg in args:
            self._build_args[arg] = None

    def get_builders(self) -> List[BuildImportArg]:
        """Get the list of BuildImportArg."""
        return list(self._build_args.keys())

    def add_from_instance(self, instance: DefaultBuilder, make_optional: bool = False) -> None:
        """
        Add the builders from another instance.

        Args:
            instance (DefaultBuilder): The instance to add the builders from.
            make_optional (bool, optional): Specifies if the import is optional. Defaults to ``False``.
        """
        for arg in instance.get_builders():
            if make_optional:
                cp = copy.copy(arg)
                cp.optional = True
                self.add_build_arg(cp)
            else:
                self.add_build_arg(arg)

    def add_ooodev_builder(
        self, mod_name: str, *, func_name: str = "get_builder", make_optional: bool = False
    ) -> None:
        """
        Add a builder from an ooodev name.

        Args:
            name (str): Ooodev name such as ``ooodev.adapter.container.index_access_partial.IndexAccessPartial``.
            optional (bool, optional): Specifies if the import is optional. Defaults to ``False``.
            make_optional (bool, optional): Specifies if the import is optional. Defaults to ``False``.
        """
        if not mod_name:
            return
        try:
            mod = self._get_import(mod_name)
        except ImportError:
            return
        if not hasattr(mod, func_name):
            return
        func = getattr(mod, func_name)
        result = func(self._component, self.lo_inst)
        self.add_from_instance(result, make_optional)

    def add_import(
        self,
        name: str,
        *,
        uno_name: str = "",
        optional: bool = False,
        init_kind: InitKind | int = InitKind.COMPONENT_INTERFACE,
        check_kind: CheckKind | int = CheckKind.NONE,
    ) -> BuildImportArg:
        """
        Add an import to the builder.

        Args:
            name (str): Ooodev name such as ``ooodev.adapter.container.index_access_partial.IndexAccessPartial``.
            uno_name (str, optional): UNO Name. such as ``com.sun.star.container.XIndexAccess``.
            optional (bool, optional): Specifies if the import is optional. Defaults to ``False``.
            init_kind (InitKind, int, optional): Init Option. Defaults to ``InitKind.COMPONENT_INTERFACE``.
            check_kind (CheckKind, int, optional): Check Kind. Defaults to ``CheckKind.NONE``.

        Returns:
            BuildImportArg: _description_

        Note:
            ``init_kind`` can be an ``InitKind`` or an ``int``:

            - NONE = 0
            - COMPONENT = 1
            - COMPONENT_INTERFACE = 2

            ``check_kind`` can be a ``CheckKind`` or an ``int``:

            - NONE = 0
            - SERVICE = 1
            - INTERFACE = 2
        """
        kind_init = InitKind(init_kind)
        kind_check = CheckKind(check_kind)
        bi = BuildImportArg(name, uno_name, optional, kind_init, kind_check)
        self.add_build_arg(bi)
        return bi

    def build_import(self, base_class: Type[Any], base_init: InitKind, name: str) -> Any:
        self._process_imports()
        clz = self._generate_class(base_class, name)
        inst = self._create_class(clz, base_init)
        self._init_classes(inst)
        return inst

    def set_omit(self, *names: str) -> None:
        """
        Set the names to omit.

        This is useful with the base class already implements the class.

        Args:
            names (str): The names to omit.

        Note:
            Names can be the name of the ``ooodev`` full import name or the UNO name.
            Valid names such as ``ooodev.adapter.container.name_access_partial.NameAccessPartial``
            or ``com.sun.star.container.XNameAccess``.
        """
        for name in names:
            self._omit.add(name)

    def get_import_names(self) -> List[str]:
        """Get the list of import names that have been added to the current instance."""
        return [arg.ooodev_name for arg in self._build_args.keys()]

    @property
    def component(self) -> Any:
        return self._component
