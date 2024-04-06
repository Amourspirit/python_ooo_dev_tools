from __future__ import annotations
import contextlib
from typing import Any, cast, Dict, List, Type, Tuple, Set, TYPE_CHECKING
import copy
import importlib
import types

import uno
from com.sun.star.lang import XServiceInfo
from com.sun.star.lang import XTypeProvider

from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.component_base import ComponentBase
from ooodev.loader import lo as mLo
from ooodev.utils.builder.build_import_arg import BuildImportArg
from ooodev.utils.builder.build_event_arg import BuildEventArg
from ooodev.utils.builder.check_kind import CheckKind
from ooodev.utils.builder.init_kind import InitKind
from ooodev.utils import gen_util as gUtil

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst


class DefaultBuilder(ComponentBase):
    def __init__(self, component: Any):
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
        self._service_info = mLo.Lo.qi(XServiceInfo, self._component, True)
        self._type_names = None

    def _get_type_names(self) -> Set[str]:
        if self._type_names is None:
            provider = mLo.Lo.qi(XTypeProvider, self._component, True)
            result: Set[str] = set()
            component_types = cast(Tuple[Any], provider.getTypes())
            for t in component_types:
                result.add(t.typeName)
            self._type_names = result
        return self._type_names

    def _get_import(self, name: str) -> types.ModuleType:
        return importlib.import_module(name)

    def _get_mod_class_names(self, name: str) -> tuple[str, str]:
        return tuple(name.rsplit(".", 1))  # type: ignore

    def _get_class(self, arg: BuildImportArg) -> Any:
        if not self._passes_check(arg):
            raise ValueError(f"Component does not support {arg.uno_name}")
        mod_name, class_name = self._get_mod_class_names(arg.ooodev_name)
        return getattr(self._get_import(mod_name), class_name)

    def _supports_interface(self, *name: str) -> bool:
        type_names = self._get_type_names()
        for n in name:
            if n in type_names:
                return True
        return False

    def _supports_only_interface(self, *name: str) -> bool:
        """
        Check if the component supports only the first interface in the list.

        Returns:
            bool: True if the component supports only the first interface in the list; otherwise, False.
        """
        count = len(name)
        if count == 0:
            return False
        type_names = self._get_type_names()
        if count == 1:
            return name[0] in type_names
        if name[0] not in type_names:
            return False
        for n in name[1:]:
            if n in type_names:
                return False
        return True

    def _supports_all_interface(self, *name: str) -> bool:
        type_names = self._get_type_names()
        for n in name:
            if n not in type_names:
                return False
        return True

    def _supports_service(self, *name: str) -> bool:
        result = False
        for srv in name:
            result = self._service_info.supportsService(srv)
            if result:
                break
        return result

    def _supports_all_service(self, *name: str) -> bool:
        result = True
        for srv in name:
            result = self._service_info.supportsService(srv)
            if not result:
                break
        return result

    def _supports_only_service(self, *name: str) -> bool:
        """
        Check if the component supports only the first services in the list.

        Returns:
            bool: True if the component supports only the first services in the list; otherwise, False.
        """
        count = len(name)
        if count == 0:
            return False
        if count == 1:
            return self._service_info.supportsService(name[0])
        result = self._service_info.supportsService(name[0])
        if not result:
            return False
        for srv in name[1:]:
            if self._service_info.supportsService(srv):
                return False
        return True

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
        elif arg.check_kind == CheckKind.SERVICE_ALL:
            return self._supports_all_service(*arg.uno_name)
        elif arg.check_kind == CheckKind.INTERFACE:
            return self._supports_interface(*arg.uno_name)
        elif arg.check_kind == CheckKind.INTERFACE_ALL:
            return self._supports_all_interface(*arg.uno_name)
        elif arg.check_kind == CheckKind.INTERFACE_ONLY:
            return self._supports_only_interface(*arg.uno_name)
        elif arg.check_kind == CheckKind.SERVICE_ONLY:
            return self._supports_only_service(*arg.uno_name)
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
        elif arg.check_kind == CheckKind.SERVICE_ALL:
            return self._supports_all_service(*arg.uno_name)
        elif arg.check_kind == CheckKind.INTERFACE:
            return self._supports_interface(*arg.uno_name)
        elif arg.check_kind == CheckKind.INTERFACE_ALL:
            return self._supports_all_interface(*arg.uno_name)
        elif arg.check_kind == CheckKind.INTERFACE_ONLY:
            return self._supports_only_interface(*arg.uno_name)
        elif arg.check_kind == CheckKind.SERVICE_ONLY:
            return self._supports_only_service(*arg.uno_name)
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
        if not name:
            raise ValueError("Name cannot be empty.")
        if name.startswith("ooodev."):
            return name
        # sample input: com.sun.star.beans.XExactName
        suffix = name.replace("com.sun.star.", "")  # beans.XExactName
        ns, class_name = suffix.rsplit(".", 1)  #  beans, XExactName
        if class_name.startswith("X"):
            class_name = class_name[1:]  # ExactName
        odev_ns = f"ooodev.adapter.{ns.lower()}.{gUtil.Util.to_snake_case(class_name)}_partial"
        odev_class = f"{class_name}Partial"
        return f"{odev_ns}.{odev_class}"

    def add_build_arg(self, *args: BuildImportArg) -> None:
        """
        Add one or more import builder to the instance.

        Args:
            args (BuildImportArg): One or more BuildImportArg instance.
        """
        for arg in args:
            self._build_args[arg] = None

    def insert_build_arg(self, idx: int, *args: BuildImportArg) -> None:
        """
        Insert one or more import builder to the instance.

        Args:
            idx (int): The index to insert the import.
            args (BuildImportArg): One or more BuildImportArg instance.
        """
        if not args:
            return
        items = list(self._build_args.items())
        if len(args) == 1:
            reversed_args = list(args)
        else:
            list_args = list(args)
            reversed_args = list_args[::-1]

        for arg in reversed_args:
            items.insert(idx, (arg, None))
        self._build_args = dict(items)

    def add_event_arg(self, *args: BuildEventArg) -> None:
        """
        Add one or more import event to the instance.

        Args:
            args (BuildEventArg): One or more BuildImportArg instance.
        """
        for arg in args:
            self._event_args[arg] = None

    def insert_event_arg(self, idx: int, *args: BuildEventArg) -> None:
        """
        Insert one or more import event to the instance.

        Args:
            idx (int): The index to insert the import.
            args (BuildEventArg): One or more BuildImportArg instance.
        """
        if not args:
            return
        items = list(self._event_args.items())
        if len(args) == 1:
            reversed_args = list(args)
        else:
            list_args = list(args)
            reversed_args = list_args[::-1]

        for arg in reversed_args:
            items.insert(idx, (arg, None))
        self._event_args = dict(items)

    def get_builders(self) -> List[BuildImportArg]:
        """Get the list of BuildImportArg."""
        return list(self._build_args.keys())

    def get_events(self) -> List[BuildEventArg]:
        """Get the list of BuildImportArg."""
        return list(self._event_args.keys())

    def merge(self, instance: DefaultBuilder, make_optional: bool = False) -> None:
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

        for arg in instance.get_events():
            if make_optional:
                self.add_event(
                    module_name=arg.module_name,
                    class_name=arg.class_name,
                    callback_name=arg.callback_name,
                    uno_name=arg.uno_name,
                    optional=True,
                    check_kind=arg.check_kind,
                )
            else:
                cp = copy.copy(arg)
                self.add_event_arg(cp)
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
        result = func(self._component)
        self.merge(result, make_optional)

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

    def insert_import(
        self,
        idx: int,
        name: str,
        *,
        uno_name: str | Tuple[str] = "",
        optional: bool = False,
        init_kind: InitKind | int = InitKind.COMPONENT_INTERFACE,
        check_kind: CheckKind | int = CheckKind.NONE,
    ) -> BuildImportArg:
        """
        Insert an import into the builder.

        Args:
            idx (int): The index to insert the import.
            name (str): Ooodev name such as ``ooodev.adapter.container.index_access_partial.IndexAccessPartial``.
            uno_name (str, optional): UNO Name. such as ``com.sun.star.container.XIndexAccess``.
            optional (bool, optional): Specifies if the import is optional. Defaults to ``False``.
            init_kind (InitKind, int, optional): Init Option. Defaults to ``InitKind.COMPONENT_INTERFACE``.
            check_kind (CheckKind, int, optional): Check Kind. Defaults to ``CheckKind.NONE``.

        Returns:
            BuildImportArg: BuildImportArg instance.
        """
        arg = self.add_import(
            name=name, uno_name=uno_name, optional=optional, init_kind=init_kind, check_kind=check_kind
        )
        _ = self._build_args.pop(arg)
        self.insert_build_arg(idx, arg)
        return arg

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
            BuildImportArg: BuildImportArg instance.

        Note:
            ``init_kind`` can be an ``InitKind`` or an ``int``:

            - NONE = 0
            - COMPONENT = 1
            - COMPONENT_INTERFACE = 2

            ``check_kind`` can be a ``CheckKind`` or an ``int``:

            - NONE = 0
            - SERVICE = 1
            - INTERFACE = 2
            - SERVICE_ALL = 3
            - INTERFACE_ALL = 4
            - SERVICE_ONLY = 5
            - INTERFACE_ONLY = 6
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

    def insert_event(
        self,
        idx: int,
        module_name: str,
        class_name: str,
        *,
        callback_name: str = "on_lazy_cb",
        uno_name: str | Tuple[str] = "",
        optional: bool = False,
        check_kind: CheckKind | int = CheckKind.INTERFACE,
    ) -> BuildEventArg:
        """
        Insert an event into the builder.

        Args:
            idx (int): The index to insert the import.
            module_name (str): Ooodev name of the module such as ``ooodev.adapter.util.refresh_events``.
            class_name (str):Ooodev class name such as ``RefreshEvents``.
            callback_name (str): Callback name such as ``on_lazy_cb``.
            uno_name (str, optional): UNO Name. such as ``com.sun.star.container.XIndexAccess``.
            optional (bool, optional): Specifies if the import is optional. Defaults to ``False``.
            check_kind (CheckKind, int, optional): Check Kind. Defaults to ``CheckKind.INTERFACE``.

        Returns:
            BuildImportArg: _description_
        """
        arg = self.add_event(
            module_name,
            class_name,
            callback_name=callback_name,
            uno_name=uno_name,
            optional=optional,
            check_kind=check_kind,
        )
        _ = self._event_args.pop(arg)
        self.insert_event_arg(idx, arg)
        return arg

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
        Add an event to the builder.

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
            - SERVICE_ALL = 3
            - INTERFACE_ALL = 4
            - SERVICE_ONLY = 5
            - INTERFACE_ONLY = 6
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
            name (str): Class Name. This can just be a class name or a full import name including the class.
                When a full import name is passed then the last part is used as the class name and parts before are used as the module name.
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
        if "." in name:
            mod, class_name = name.rsplit(".", 1)
        else:
            mod = ""
            class_name = name
        clz = self._generate_class(base_class, class_name)
        if mod:
            with contextlib.suppress(Exception):
                clz.__module__ = mod
        inst = self._create_class(clz, InitKind(init_kind))
        self._init_classes(inst)
        return inst

    def set_omit(self, *names: str) -> None:
        """
        Set the names to omit.

        This is useful with the base class already implements the class.

        Args:
            names (str): The CASE Sensitive names to omit.

        Note:
            Names can be the name of the ``ooodev`` full import name or the UNO name.
            Valid names such as ``ooodev.adapter.container.name_access_partial.NameAccessPartial``
            or ``com.sun.star.container.XNameAccess``.
        """
        for name in names:
            self._omit.add(self._convert_to_ooodev(name))

    def get_import_names(self) -> List[str]:
        """Get the list of import names that have been added to the current instance."""
        return [arg.ooodev_name for arg in self._build_args.keys()]

    @property
    def component(self) -> Any:
        return self._component

    @property
    def omits(self) -> Set[str]:
        return self._omit
