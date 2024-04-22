from __future__ import annotations
from typing import cast, Union, TYPE_CHECKING
from ooodev.gui.menu.popup.popup_item import PopupItem
from ooodev.gui.menu.popup_menu import PopupMenu
from ooodev.gui.menu.shortcuts import Shortcuts
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.args.event_args import EventArgs
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.gui.commands.cmd_info import CmdInfo
from ooodev.utils.kind.module_names_kind import ModuleNamesKind


if TYPE_CHECKING:
    from ooodev.gui.menu.common.command_dict import CommandDict


class MenuProcessor(EventsPartial):
    """Class for processing menus. Does not process submenus"""

    def __init__(self, popup: PopupMenu, cmd_info: CmdInfo) -> None:
        """
        Constructor

        Args:
            item (PopupItem): Menu item data.
        """
        EventsPartial.__init__(self)
        self._popup = popup
        self._cmd_info = cmd_info

    def _process_shortcut(self, pop: PopupItem) -> None:
        """Process shortcut"""
        keys = pop.shortcut.strip()
        if not keys:
            return
        kv = Shortcuts.to_key_event(keys)
        if kv is not None:
            self._popup.set_accelerator_key_event(pop.menu_id, kv)

    def _process_command(self, pop: PopupItem) -> None:
        """Process command"""
        if not pop.command:
            return
        cmd = Shortcuts.get_url_script(pop.command)
        self._popup.set_command(pop.menu_id, cmd)

    def _get_popup_from_module(self, menu: dict, index: int) -> PopupItem:
        module = cast(str, menu["module"])
        module_kind = ModuleNamesKind(module)
        cmd = cast(str, menu["command"])
        cmd_info = self._cmd_info.get_cmd_data(module_kind, cmd)
        if cmd_info is None:
            raise ValueError(f"Command not found: {module} {cmd}")
        menu_data = {
            "command": cmd,
            "text": cmd_info.label or cmd_info.name or cmd,
            "tip_help_text": cmd_info.tooltip_label,
        }
        menu_data.update(menu)
        return self._get_popup_item(menu_data, index)

    def _get_popup_item(self, menu: dict, index: int) -> PopupItem:
        """Get popup item data"""
        eargs = EventArgs(self)

        eargs.event_data = {**menu}
        eargs.event_data["index"] = index
        self.trigger_event("before_get_popup_item", eargs)
        menu_data = cast(dict, eargs.event_data)
        text = menu_data["text"]
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
            command=command,
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

    def process(self, menu: dict, index: int) -> PopupItem:
        """Process menu item"""
        if "module" in menu:
            pop = self._get_popup_from_module(menu, index)
        else:
            pop = self._get_popup_item(menu, index)
        cargs = CancelEventArgs(source=self)
        cargs.event_data = {"popup_menu": self._popup, "popup_item": pop}
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
