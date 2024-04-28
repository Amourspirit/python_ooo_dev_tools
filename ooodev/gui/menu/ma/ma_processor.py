from __future__ import annotations
from typing import cast, Any, List, Union, TYPE_CHECKING
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.gui.commands.cmd_data import CmdData
from ooodev.gui.commands.cmd_info import CmdInfo
from ooodev.gui.menu.shortcuts import Shortcuts
from ooodev.utils.helper.dot_dict import DotDict
from ooodev.utils.kind.module_names_kind import ModuleNamesKind
from ooodev.gui.menu.ma.ma_item import MAItem

if TYPE_CHECKING:
    from ooodev.gui.menu.common.command_dict import CommandDict


class MAProcessor(EventsPartial):
    """Class for processing App menus. Does not process submenus"""

    def __init__(self, container: List[MAItem]) -> None:
        """
        Constructor

        Args:
            container (List[MAItem]): Menu item data.
        """
        EventsPartial.__init__(self)
        self._init_events()
        self._container = container
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

    def _process_command(self, pop: MAItem) -> None:
        """Process command"""
        if not pop.command:
            return
        if pop.is_separator():
            return
        if isinstance(pop.command, str):
            cmd = pop.command
        else:
            cmd = pop.command.to_dict()
        pop.command = Shortcuts.get_url_script(cmd)

    def _get_action_from_module(self, menu: dict, index: int) -> MAItem | None:
        """
        Gets the Action Trigger item and looks up the command data from the command info.

        Args:
            menu (dict): Menu Data
            index (int): Current Index

        Raises:
            ValueError: If no text is found for the command

        Returns:
            MAItem | None: Action Item if found, otherwise None

        Note:
            If the menu text is not found then the event ``ma_item_module_no_text_found`` is raised.
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

        module_kind = ModuleNamesKind(menu["module"])
        cmd = cast(str, menu["CommandURL"])
        command_data = _get_command_data(module_kind, cmd)
        if command_data is None:
            cmd_text = ""
        else:
            cmd_text = get_cmd_text(command_data)

        if not cmd_text:
            # check if a label has bee provided:
            if "Label" in menu:
                cmd_text = menu["Label"]
            else:
                # no valid menu text found
                # raise an event to allow the caller to provide a valid menu text
                cargs = CancelEventArgs(self)
                event_data = DotDict(module_kind=module_kind, cmd=cmd, index=index, menu=menu)
                cargs.event_data = event_data
                self.trigger_event("ma_item_module_no_text_found", cargs)
                if cargs.cancel:
                    return None
                if "Label" in cargs.event_data.menu:
                    cmd_text = cargs.event_data.menu["Label"]

        if not cmd_text:
            raise ValueError(f"No text for command: {module_kind} {cmd}")

        menu_data = {
            "CommandURL": cmd,
        }
        menu_data.update(menu)
        # set the text back to the found command text
        menu_data["Label"] = cmd_text
        return self._get_action_item(menu_data, index)

    def _get_action_item(self, menu: dict, index: int) -> MAItem:
        """Get popup item data"""
        eargs = EventArgs(self)

        eargs.event_data = DotDict(**menu)
        eargs.event_data.index = index
        self.trigger_event("before_get_action_item", eargs)
        menu_data = cast(DotDict, eargs.event_data)
        text = menu_data.Label
        command = cast(Union[str, "CommandDict"], menu_data.get("CommandURL", ""))
        return MAItem(
            label=text,
            command=Shortcuts.get_url_script(command),
            style=menu_data.get("Style", 0),
            shortcut=menu_data.get("ShortCut", ""),
            submenu=menu_data.get("Submenu"),
        )

    def get_action_item(self, menu: dict, index: int) -> MAItem | None:
        """Get Action item"""
        if "module" in menu:
            item = self._get_action_from_module(menu, index)
        elif "Label" in menu:
            text = menu["Label"]
            if text == "-":
                item = MAItem(label="-")
            else:
                item = self._get_action_item(menu, index)
        else:
            item = self._get_action_item(menu, index)
        return item

    def process(self, menu: dict, index: int) -> MAItem | None:
        """Process menu item"""
        item = self.get_action_item(menu, index)
        if item is None:
            return None
        cargs = CancelEventArgs(source=self)
        cargs.event_data = DotDict(container=self._container, action_item=item)
        self.trigger_event("before_process", cargs)
        if cargs.cancel:
            return item

        self._container.insert(index, item)
        if item.is_separator():
            return item

        self._process_command(item)
        eargs = EventArgs.from_args(cargs)
        self.trigger_event("after_process", eargs)
        return item
