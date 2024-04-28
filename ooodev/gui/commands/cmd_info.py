from __future__ import annotations
from datetime import timedelta
from typing import Any, Callable, Dict, List, Tuple, TYPE_CHECKING

from ooodev.adapter.frame.module_manager_comp import ModuleManagerComp
from ooodev.adapter.frame.the_ui_command_description_comp import TheUICommandDescriptionComp
from ooodev.adapter.ui.accelerator_configuration_comp import AcceleratorConfigurationComp
from ooodev.adapter.ui.global_accelerator_configuration_comp import GlobalAcceleratorConfigurationComp
from ooodev.adapter.ui.the_module_ui_configuration_manager_supplier_comp import (
    TheModuleUIConfigurationManagerSupplierComp,
)
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.gui.commands.cmd_data import CmdData
from ooodev.utils import props as mProps
from ooodev.utils.cache.file_cache.pickle_cache import PickleCache
from ooodev.utils.cache.time_cache import TimeCache
from ooodev.utils.kind.module_names_kind import ModuleNamesKind
from ooodev.utils.string.str_list import StrList

if TYPE_CHECKING:
    from ooodev.gui.menu.shortcuts import Shortcuts


class CmdInfo(EventsPartial):
    """
    Singleton Class. Gets Information about commands.

    Processing all the commands in all the modules takes some time.
    The first time the command data is retrieved, it is cached.
    The next time the command data is retrieved, it is retrieved from the cache.
    The cache valid for 5 days (may be cleared sooner, depending on system).
    The ``clear_cache()`` method can be used to clear the cache sooner.

    Example:
        .. code-block:: python

            >>> from ooodev.utils.kind.module_names_kind import ModuleNamesKind
            >>> from ooodev.gui.commands.cmd_info import CmdInfo

            >>> inst = CmdInfo()
            >>> cmd_data = inst.get_cmd_data(ModuleNamesKind.SPREADSHEET_DOCUMENT, ".uno:Copy")
            >>> if cmd_data:
            ...     print(cmd_data)
            CmdData(
                command='.uno:Copy',
                label='Cop~y',
                name='Copy',
                popup=False,
                properties=1,
                popup_label='',
                tooltip_label='',
                target_url='',
                is_experimental=False,
                module_hotkey='',
                global_hotkey='Ctrl+C'
            )

    See Also:
        - :ref:`help_getting_info_on_commands`
    """

    _instance = None

    def __new__(cls, *args, **kwargs):
        if not isinstance(cls._instance, cls):
            # avoid circular import
            # pylint: disable=import-outside-toplevel
            from ooodev.gui.menu.shortcuts import Shortcuts

            cls._instance = super(CmdInfo, cls).__new__(cls, *args, **kwargs)
            cls._instance._initialized = False
            cls._instance._global_shortcuts = Shortcuts()
        return cls._instance

    # see https://opengrok.libreoffice.org/xref/core/officecfg/registry/data/org/openoffice/Office/UI/
    def __init__(self) -> None:
        """Constructor."""
        if self._initialized:
            return
        EventsPartial.__init__(self)
        self._cmf = TheModuleUIConfigurationManagerSupplierComp.from_lo()
        self._ui_cmd_desc = TheUICommandDescriptionComp.from_lo()
        delta = timedelta(days=5)
        self._file_cache = PickleCache(tmp_dir="ooodev/gui/commands_cmd_info", lifetime=delta.total_seconds())
        # using mem cache is many times faster than using file cache alone.
        self._mem_cache = TimeCache(300, 300)
        self._file_prefix = "uurt54_cmds_"
        self._global_shortcuts: Shortcuts
        self._initialized = True

    def get_module_names(self) -> Tuple[str, ...]:
        """
        Get a list of module names such as ``com.sun.star.presentation.PresentationDocument``.
        """
        key = "get_module_names"
        if key in self._mem_cache:
            return self._mem_cache[key]
        mm = ModuleManagerComp.from_lo()
        names = mm.get_element_names()
        self._mem_cache[key] = names
        return names

    def clear_cache(self) -> None:
        """Clear the cache."""
        self._mem_cache.clear()
        names = self.get_module_names()
        for name in names:
            key = f"{self._file_prefix}{name.replace('.', '_')}.pkl"
            self._file_cache.remove(key)
        self._mem_cache.clear()

    def find_command(self, command: str) -> Dict[str, List[CmdData]]:
        """
        Find a command by name. All Modules are searched and the dictionary key is the module name.

        Args:
            command (str): The command name such as ``.uno:Copy``.

        Returns:
            dict: A dictionary of module names and command rows.

        Note:
            Triggers the ``CmdInfo_command_found`` event. If the event is canceled the command is not added to the result.
            See :py:meth:`~.CmdInfo.subscribe_on_command_found`.

            The ``event_data`` has the following event data:
            - ``module_name``: The module name.
            - ``cmd_data``: Instance of :py:class:`~ooodev.gui.commands.CmdData` The command data.
        """
        results = {}
        for name in self.get_module_names():
            cmd_data = self.get_cmd_data(name, command)
            if cmd_data:
                cargs = CancelEventArgs(self)
                cargs.event_data = {"module_name": name, "cmd_data": cmd_data}
                self.trigger_event("CmdInfo_command_found", cargs)
                if cargs.cancel:
                    continue
                if name not in results:
                    results[name] = [cmd_data]
                else:
                    results[name].append(cmd_data)
        return results

    def has_command(self, mode_name: str | ModuleNamesKind, cmd: str) -> bool:
        """
        Gets if a module contains a command.

        Args:
            mode_name (str | ModuleNamesKind): Module name such as ``com.sun.star.presentation.PresentationDocument``.
            cmd (str): Command name such as ``.uno:Copy``.

        Returns:
            bool: ``True`` if the command is found; Otherwise, ``False``.

        Hint:
            - ``ModuleNamesKind`` is an enum and can be imported from ``ooodev.utils.kind.module_names_kind``.
        """
        return self.get_cmd_data(mode_name, cmd) is not None

    def get_cmd_data(self, mode_name: str | ModuleNamesKind, cmd: str) -> CmdData | None:
        """
        Gets the command data.

        Args:
            mode_name (str | ModuleNamesKind): Module name such as ``com.sun.star.presentation.PresentationDocument``.
            cmd (str): Command name such as ``.uno:Copy``.

        Returns:
            CmdData | None: The command data or ``None`` if not found.

        Hint:
            - ``ModuleNamesKind`` is an enum and can be imported from ``ooodev.utils.kind.module_names_kind``.
        """
        dic = self.get_dict(str(mode_name))
        row = dic.get(cmd)
        if row:
            return CmdData(*row)
        return None

    def get_dict(self, mod_name: str | ModuleNamesKind) -> Dict[str, CmdData]:
        """
        Gets a dictionary of command data. The key is the command name such as ``.uno:Copy``.

        Args:
            mod_name (str | ModuleNamesKind): Module name such as ``com.sun.star.presentation.PresentationDocument``.

        Returns:
            Dict[str, CmdData]: A dictionary of command data.

        Hint:
            - ``ModuleNamesKind`` is an enum and can be imported from ``ooodev.utils.kind.module_names_kind``.
        """
        # names = self.get_module_names()
        mod_name = str(mod_name)
        key = f"{self._file_prefix}{mod_name.replace('.', '_')}.pkl"
        if key in self._mem_cache:
            return self._mem_cache[key]
        val = self._file_cache[key]
        if val:
            self._mem_cache[key] = val
            return val

        # config = self._cmf.get_ui_configuration_manager(mod_name)
        def build_row(
            cmd: str, el: Any, local_sc: AcceleratorConfigurationComp, global_sc: GlobalAcceleratorConfigurationComp
        ) -> tuple:
            el_dict = mProps.Props.data_to_dict(el)
            label = el_dict.get("Label", "")
            name = el_dict.get("Name", "")
            popup = el_dict.get("Popup", False)
            properties = el_dict.get("Properties", 0)
            popup_label = el_dict.get("PopupLabel", "")
            tooltip_label = el_dict.get("TooltipLabel", "")
            target_url = el_dict.get("TargetURL", "")
            is_experimental = el_dict.get("IsExperimental", False)
            local_keys = local_sc.get_preferred_key_events_for_command_list(cmd)
            global_keys = global_sc.get_preferred_key_events_for_command_list(cmd)
            if local_keys:
                sl = StrList()
                for key in local_keys:
                    sl.append(self._global_shortcuts.from_key_event(key))
                module_hotkey = str(sl)
            else:
                module_hotkey = ""
            if global_keys:
                sl = StrList()
                for key in global_keys:
                    sl.append(self._global_shortcuts.from_key_event(key))
                global_hotkey = str(sl)
            else:
                global_hotkey = ""

            return CmdData(
                cmd,
                label,
                name,
                popup,
                properties,
                popup_label,
                tooltip_label,
                target_url,
                is_experimental,
                module_hotkey,
                global_hotkey,
            )

        result = {}
        desc = self._ui_cmd_desc.get_by_name(mod_name)
        commands = desc.get_element_names()
        mod_ui_config = self._cmf.get_ui_configuration_manager(mod_name)
        local_short_cut_mgr = mod_ui_config.get_short_cut_manager()
        global_short_cut_mgr = GlobalAcceleratorConfigurationComp.from_lo()

        for name in commands:
            el = desc.get_by_name(name)
            row = build_row(name, el, local_short_cut_mgr, global_short_cut_mgr)
            result[name] = row
        self._file_cache[key] = result
        self._mem_cache[key] = result
        return result

    def subscribe_on_command_found(self, cb: Callable[[Any, CancelEventArgs], None]) -> None:
        """
        Subscribe on command found event.

        The callback ``event_data`` is a dictionary with keys:

        - ``module_name``: The module name.
        - ``cmd_data``: Instance of :py:class:`~ooodev.gui.commands.CmdData` The command data.
        """
        self.subscribe_event("CmdInfo_command_found", cb)

    def unsubscribe_on_command_found(self, cb: Callable[[Any, CancelEventArgs], None]) -> None:
        """Unsubscribe on command found event."""
        self.unsubscribe_event("CmdInfo_command_found", cb)
