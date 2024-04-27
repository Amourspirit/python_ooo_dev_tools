from __future__ import annotations
from typing import cast, Any, Union, TYPE_CHECKING
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.gui.commands.cmd_data import CmdData
from ooodev.gui.commands.cmd_info import CmdInfo
from ooodev.gui.menu.popup_menu import PopupMenu
from ooodev.gui.menu.popup.popup_item import PopupItem
from ooodev.gui.menu.shortcuts import Shortcuts
from ooodev.utils.helper.dot_dict import DotDict
from ooodev.utils.kind.module_names_kind import ModuleNamesKind
from ooodev.utils.string.str_list import StrList

if TYPE_CHECKING:
    from ooodev.gui.menu.common.command_dict import CommandDict


class PopupProcessor(EventsPartial):
    """
    Class for processing menus. Does not process submenus.

    See Also:
        - :ref:`help_popup_from_dict_or_json`
        - :ref:`help_popup_via_builder_item`
    """

    def __init__(self, popup: PopupMenu) -> None:
        """
        Constructor

        Args:
            item (PopupItem): Menu item data.
        """
        EventsPartial.__init__(self)
        self._init_events()
        self._popup = popup
        self._cmd_info = CmdInfo()
        self._cmd_info.subscribe_on_command_found(self._fn_on_command_found)

    def _init_events(self) -> None:
        self._fn_on_command_found = self._on_command_found

    def _on_command_found(self, source: Any, event: CancelEventArgs) -> None:
        """Command found event"""
        pass
        # cmd_data = cast(CmdData, event.event_data["cmd_data"])
        # if not cmd_data.global_hotkey:
        #     event.cancel = True

    def _process_shortcut(self, pop: PopupItem) -> None:
        """Process shortcut"""
        keys = pop.shortcut.strip()
        if not keys:
            return
        sl_keys = StrList.from_str(keys)
        for key in sl_keys:
            if not key:
                continue
            kv = Shortcuts.to_key_event(key)
            if kv is not None:
                self._popup.set_accelerator_key_event(pop.menu_id, kv)

    def _process_command(self, pop: PopupItem) -> None:
        """Process command"""
        if not pop.command:
            return
        cmd = Shortcuts.get_url_script(pop.command)
        self._popup.set_command(pop.menu_id, cmd)

    def _get_popup_from_module(self, menu: dict, index: int) -> PopupItem | None:
        """
        Gets the popup item and looks up the command data from the command info.

        Args:
            menu (dict): Menu Data
            index (int): Current Index

        Raises:
            ValueError: If no text is found for the command

        Returns:
            PopupItem | None: Popup Item if found, otherwise None

        Note:
            If the menu text is not found then the event ``popup_module_no_text_found`` is raised.
            The event data is a dictionary with keys:
            - ``module_kind``: ModuleNamesKind
            - ``cmd``: Command as a string.
            - ``index``: Index as an integer.
            - ``menu``: Menu Data as a dictionary.

            The caller can set ``menu["text"]`` to provide a valid menu text.
            If the caller cancels the event then the menu item is not created and None is returned.
        """

        def _get_command_data(kind: ModuleNamesKind, cmd: str) -> CmdData | None:
            if kind == ModuleNamesKind.NONE:
                # get from global
                data = self._cmd_info.find_command(cmd)
                if cmd not in data:
                    return None
                c_infos = data[cmd]
                if not c_infos:
                    return None
                return c_infos[0]
            return self._cmd_info.get_cmd_data(kind, cmd)

        def get_cmd_text(data: CmdData) -> str:
            return data.label or data.name

        def get_shortcut(data: CmdData) -> str:
            if data.module_hotkey:
                return data.module_hotkey
            if data.global_hotkey:
                # str_keys = StrList.from_str(data.global_hotkey)
                return data.global_hotkey
            return ""

        module_kind = ModuleNamesKind(menu["module"])
        cmd = cast(str, menu["command"])
        command_data = _get_command_data(module_kind, cmd)
        if command_data is None:
            cmd_text = ""
        else:
            cmd_text = get_cmd_text(command_data)

        if not cmd_text:
            # check if a label has bee provided:
            if "text" in menu:
                cmd_text = menu["text"]
            else:
                # no valid menu text found
                # raise an event to allow the caller to provide a valid menu text
                cargs = CancelEventArgs(self)
                event_data = DotDict(module_kind=module_kind, cmd=cmd, index=index, menu=menu)
                cargs.event_data = event_data
                self.trigger_event("popup_module_no_text_found", cargs)
                if cargs.cancel:
                    return None
                if "text" in cargs.event_data.menu:
                    cmd_text = cargs.event_data.menu["text"]

        if not cmd_text:
            raise ValueError(f"No text for command: {module_kind} {cmd}")

        menu_data = {
            "command": cmd,
        }
        if command_data is not None and command_data.tooltip_label:
            menu_data["tip_help_text"] = command_data.tooltip_label
        menu_data.update(menu)
        # set the text back to the found command text
        menu_data["text"] = cmd_text
        if command_data is not None:
            if "shortcut" not in menu_data:
                shortcut = get_shortcut(command_data)
                if shortcut:
                    menu_data["shortcut"] = shortcut
        return self._get_popup_item(menu_data, index)

    def _get_popup_item(self, menu: dict, index: int) -> PopupItem:
        """Get popup item data"""
        eargs = EventArgs(self)

        eargs.event_data = DotDict(**menu)
        eargs.event_data.index = index
        self.trigger_event("before_get_popup_item", eargs)
        menu_data = cast(DotDict, eargs.event_data)
        text = menu_data.text
        command = cast(Union[str, "CommandDict"], menu_data.get("command", ""))
        style = int(menu_data.get("style", 0))
        checked = bool(menu_data.get("checked", False))
        enabled = bool(menu_data.get("enabled", True))
        default = bool(menu_data.get("default", False))
        help_command = menu_data.get("help_command", "")
        help_text = menu_data.get("help_text", "")
        tip_help_text = menu_data.get("tip_help_text", "")
        shortcut = cast(str, menu_data.get("shortcut", ""))
        data = menu_data.get("data")
        menu_id = int(menu_data.get("menu_id", index))
        return PopupItem(
            text=text,
            menu_id=menu_id,
            command=Shortcuts.get_url_script(command),
            style=style,
            checked=checked,
            enabled=enabled,
            default=default,
            help_command=help_command,
            help_text=help_text,
            tip_help_text=tip_help_text,
            shortcut=shortcut,
            index=index,
            data=data,
        )

    def get_popup_item(self, menu: dict, index: int) -> PopupItem | None:
        """Get popup item"""
        if "module" in menu:
            pop = self._get_popup_from_module(menu, index)
        else:
            pop = self._get_popup_item(menu, index)
        return pop

    def process(self, menu: dict, index: int) -> PopupItem | None:
        """Process menu item"""
        pop = self.get_popup_item(menu, index)
        if pop is None:
            return None
        cargs = CancelEventArgs(source=self)
        cargs.event_data = DotDict(popup_menu=self._popup, popup_item=pop)
        self.trigger_event("before_process", cargs)
        if cargs.cancel:
            return pop

        if pop.is_separator():
            self._popup.insert_separator(pop.index)
            return pop

        self._popup.insert_item(pop.menu_id, pop.text, pop.style, pop.index)
        self._process_command(pop)
        self._popup.enable_item(pop.menu_id, pop.enabled)
        if pop.checked:
            self._popup.check_item(pop.menu_id)
        if pop.default:
            self._popup.set_default_item(pop.menu_id)
        if pop.help_command:
            self._popup.set_help_command(pop.menu_id, pop.help_command)
        if pop.help_text:
            self._popup.set_help_text(pop.menu_id, pop.help_text)
        if pop.tip_help_text:
            self._popup.set_tip_help_text(pop.menu_id, pop.tip_help_text)
        if pop.shortcut:
            self._process_shortcut(pop)
        eargs = EventArgs.from_args(cargs)
        self.trigger_event("after_process", eargs)
        return pop
