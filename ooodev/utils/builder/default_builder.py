from __future__ import annotations
from typing import Any, Dict, List, Type, Tuple, Set, TYPE_CHECKING
import copy
import importlib
import types

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.component_base import ComponentBase
from ooodev.loader import lo as mLo
from ooodev.utils import info as mInfo
from ooodev.utils.builder.build_import_arg import BuildImportArg
from ooodev.utils.builder.build_event_arg import BuildEventArg
from ooodev.utils.builder.check_kind import CheckKind
from ooodev.utils.builder.init_kind import InitKind
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils import gen_util as gUtil

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
        self._bases_partial: Dict[Type[Any], BuildImportArg] = {}
        # _bases_interfaces could be interfaces such as ComponentPartial
        self._bases_interfaces: Dict[Type[Any], BuildImportArg] = {}
        self._bases_event_interfaces: Dict[types.ModuleType, BuildEventArg] = {}
        # _build_args is a dictionary of BuildImportArg, using dictionary to avoid duplicates and to keep order
        self._build_args: Dict[BuildImportArg, Any] = {}
        self._event_args: Dict[BuildEventArg, Any] = {}
        self._omit: Set[str] = set()

    def _get_import(self, name: str) -> types.ModuleType:
        return importlib.import_module(name)

    def _get_mod_class_names(self, name: str) -> tuple[str, str]:
        return tuple(name.rsplit(".", 1))  # type: ignore

    def _get_class(self, arg: BuildImportArg) -> Any:
        if not self._passes_check(arg):
            raise ValueError(f"Component does not support {arg.uno_name}")
        mod_name, class_name = self._get_mod_class_names(arg.ooodev_name)
        return getattr(self._get_import(mod_name), class_name)

    def _has_interface(self, *name: str) -> bool:
        return mLo.Lo.is_uno_interfaces(self._component, *name)

    def _supports_service(self, *name: str) -> bool:
        return mInfo.Info.support_service(self._component, *name)

    def _get_optional_class(self, arg: BuildImportArg) -> Any:
        if not arg.uno_name:
            return None
        if self._passes_check(arg):
            return self._get_class(arg)
        return None

    def _get_optional_event(self, arg: BuildEventArg) -> types.ModuleType | None:
        if self._passes_event_check(arg):
            return self._get_import(arg.module_name)
        return None

    def _passes_check(self, arg: BuildImportArg) -> bool:
        if not arg.uno_name:
            return True
        if arg.check_kind == CheckKind.SERVICE:
            return self._supports_service(*arg.uno_name)
        elif arg.check_kind == CheckKind.INTERFACE:
            return self._has_interface(*arg.uno_name)
        return True

    def _passes_event_check(self, arg: BuildEventArg) -> bool:
        if not arg.module_name:
            return False
        if not arg.class_name:
            return False
        if not arg.callback_name:
            return False
        if arg.check_kind == CheckKind.SERVICE:
            return self._supports_service(*arg.uno_name)
        elif arg.check_kind == CheckKind.INTERFACE:
            return self._has_interface(*arg.uno_name)
        return True

    def _add_base(self, base: Type[Any], arg: BuildImportArg) -> None:
        if arg.init_kind == InitKind.COMPONENT:
            if self._passes_check(arg):
                self._bases_partial[base] = arg
        else:
            if self._passes_check(arg):
                self._bases_interfaces[base] = arg

    def _add_event_base(self, mod_type: types.ModuleType, arg: BuildEventArg) -> None:
        self._bases_event_interfaces[mod_type] = arg

    def _clear_imports(self) -> None:
        self._bases_partial.clear()
        self._bases_interfaces.clear()
        self._bases_event_interfaces.clear()

    def _process_imports(self) -> None:
        self._clear_imports()
        for arg in self._build_args.keys():
            if not arg.ooodev_name:
                continue
            if arg.ooodev_name in self._omit:
                continue
            if arg.optional:
                clz = self._get_optional_class(arg)
                if clz:
                    self._add_base(clz, arg)
            else:
                clz = self._get_class(arg)
                self._add_base(clz, arg)

        for arg in self._event_args.keys():
            if not arg.module_name:
                continue
            if not arg.class_name:
                continue
            if not arg.callback_name:
                continue
            if arg.optional:
                mod = self._get_optional_event(arg)
                if mod:
                    self._add_event_base(mod, arg)
            else:
                mod = self._get_import(arg.module_name)
                self._add_event_base(mod, arg)

    def _init_class(self, instance: Any, clz: Type[Any], kind: InitKind) -> Any:
        if kind == InitKind.COMPONENT:
            clz.__init__(instance, self._component)  # type: ignore
        elif kind == InitKind.COMPONENT_INTERFACE:
            clz.__init__(instance, self._component, None)  # type: ignore
        else:
            clz.__init__(instance)

    def _init_event_class(self, instance: Any, mod: types.ModuleType, arg: BuildEventArg) -> Any:
        # after instance is created, we need to callback to the module on_lazy_cb()
        clz = getattr(mod, arg.class_name)
        trigger_args = GenericArgs(src_comp=self._component, src_instance=instance)
        cb = getattr(mod, arg.callback_name)
        # obj = clz.__new__(clz)  # type: ignore
        # obj.__init__(trigger_args=trigger_args, cb=cb)
        clz.__init__(instance, trigger_args=trigger_args, cb=cb)  # type: ignore

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
            self._init_class(instance, clz, kind.init_kind)
        for clz, kind in self._bases_interfaces.items():
            self._init_class(instance, clz, kind.init_kind)
        for mod, arg in self._bases_event_interfaces.items():
            self._init_event_class(instance, mod, arg)

    def _generate_class(self, base_class: Type[Any], name: str, **kwargs) -> Type[Any]:
        bases = (
            [base_class]
            + list(self._bases_partial.keys())
            + list(self._bases_interfaces.keys())
            + list([getattr(mod, arg.class_name) for mod, arg in self._bases_event_interfaces.items()])
        )
        if len(bases) == 1 and not kwargs:
            return base_class
        else:
            return type(name, tuple(bases), kwargs)

    def _convert_to_ooodev(self, name: str) -> str:
        # sample input: com.sun.star.beans.XExactName
        suffix = name.replace("com.sun.star.", "")  # beans.XExactName
        ns, class_name = suffix.rsplit(".", 1)  #  beans, XExactName
        if class_name.startswith("X"):
            class_name = class_name[1:]  # ExactName
        odev_ns = f"ooodev.adapter.{ns.lower()}.{gUtil.Util.to_snake_case(class_name)}_partial"
        odev_class = f"{class_name}Partial"
        return f"{odev_ns}.{odev_class}"

    def add_build_arg(self, *args: BuildImportArg) -> None:
        for arg in args:
            self._build_args[arg] = None

    def add_event_arg(self, *args: BuildEventArg) -> None:
        for arg in args:
            self._event_args[arg] = None

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
                self.add_import(
                    name=arg.ooodev_name,
                    uno_name=arg.uno_name,
                    optional=True,
                    init_kind=arg.init_kind,
                    check_kind=arg.check_kind,
                )
            else:
                cp = copy.copy(arg)
                self.add_build_arg(cp)
        self.omits.update(instance.omits)

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

    def remove_import(self, name: str) -> None:
        """
        Remove an import from the builder.

        Args:
            name (str): Ooodev name such as ``ooodev.adapter.container.index_access_partial.IndexAccessPartial``.
        """
        for arg in self._build_args.keys():
            if arg.ooodev_name == name:
                del self._build_args[arg]
                break

    def remove_event(self, module_name: str) -> None:
        """
        Remove an event from the builder.

        Args:
            module_name (str): Ooodev name such as ``ooodev.adapter.util.refresh_events.RefreshEvents``.
        """
        for arg in self._event_args.keys():
            if arg.module_name == module_name:
                del self._event_args[arg]
                break

    def auto_add_interface(self, uno_name: str, optional: bool = True) -> None:
        """
        Add an import from a UNO name.

        Args:
            uno_name (str): UNO Name. such as ``com.sun.star.container.XIndexAccess``.
        """
        name = self._convert_to_ooodev(uno_name)
        self.add_import(name=name, uno_name=uno_name, optional=optional, check_kind=CheckKind.INTERFACE)

    def add_import(
        self,
        name: str,
        *,
        uno_name: str | Tuple[str] = "",
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
        if isinstance(uno_name, str):
            uname = uno_name.strip()
            if uname:
                uno_name = (uname,)
            else:
                uno_name = ()  # type: ignore
        kind_init = InitKind(init_kind)
        kind_check = CheckKind(check_kind)
        bi = BuildImportArg(
            ooodev_name=name.strip(),
            uno_name=uno_name,  # type: ignore
            optional=optional,
            init_kind=kind_init,
            check_kind=kind_check,
        )
        self.add_build_arg(bi)
        return bi

    def add_event(
        self,
        module_name: str,
        class_name: str,
        *,
        callback_name: str = "on_lazy_cb",
        uno_name: str | Tuple[str] = "",
        optional: bool = False,
        check_kind: CheckKind | int = CheckKind.INTERFACE,
    ) -> BuildEventArg:
        """
        Add an import to the builder.

        Args:
            module_name (str): Ooodev name of the module such as ``ooodev.adapter.util.refresh_events``.
            class_name (str):Ooodev class name such as ``RefreshEvents``.
            callback_name (str): Callback name such as ``on_lazy_cb``.
            uno_name (str, optional): UNO Name. such as ``com.sun.star.container.XIndexAccess``.
            optional (bool, optional): Specifies if the import is optional. Defaults to ``False``.
            check_kind (CheckKind, int, optional): Check Kind. Defaults to ``CheckKind.INTERFACE``.

        Returns:
            BuildImportArg: _description_

        Note:
            ``check_kind`` can be a ``CheckKind`` or an ``int``:

            - NONE = 0
            - SERVICE = 1
            - INTERFACE = 2
        """
        if isinstance(uno_name, str):
            uname = uno_name.strip()
            if uname:
                uno_name = (uname,)
            else:
                uno_name = ()  # type: ignore
        kind_check = CheckKind(check_kind)
        bi = BuildEventArg(
            module_name=module_name.strip(),
            class_name=class_name.strip(),
            callback_name=callback_name.strip(),
            uno_name=uno_name,  # type: ignore
            optional=optional,
            check_kind=kind_check,
        )
        self.add_event_arg(bi)
        return bi

    def build_class(self, name: str, base_class: Type[Any], init_kind: InitKind | int = InitKind.COMPONENT) -> Any:
        """
        Build the import.

        Args:
            name (str): Class Name
            init_kind (InitKind, int, optional): Init Option. Defaults to ``InitKind.COMPONENT``.
            base_class (Type[Any], optional): Base Class. Defaults to ``BuilderBase``.

        Returns:
            Any: Class instance

        Note:
            ``init_kind`` can be an ``InitKind`` or an ``int``:

            - NONE = 0
            - COMPONENT = 1
            - COMPONENT_INTERFACE = 2
        """
        self._process_imports()
        clz = self._generate_class(base_class, name)
        inst = self._create_class(clz, InitKind(init_kind))
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

    @property
    def omits(self) -> Set[str]:
        return self._omit
