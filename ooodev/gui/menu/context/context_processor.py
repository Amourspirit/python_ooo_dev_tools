from __future__ import annotations
from typing import cast, Any, Union, TYPE_CHECKING
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.gui.commands.cmd_data import CmdData
from ooodev.gui.commands.cmd_info import CmdInfo
from ooodev.gui.menu.context.action_trigger_item import ActionTriggerItem
from ooodev.gui.menu.context.action_trigger_sep import ActionTriggerSep
from ooodev.gui.menu.context.action_trigger_container import ActionTriggerContainer
from ooodev.gui.menu.shortcuts import Shortcuts
from ooodev.utils.helper.dot_dict import DotDict
from ooodev.utils.kind.module_names_kind import ModuleNamesKind

if TYPE_CHECKING:
    from ooodev.gui.menu.common.command_dict import CommandDict


class ContextProcessor(EventsPartial):
    """
    Class for processing context menus for interception. Does not process submenus.

    See Also:
        - :ref:`help_menu_context_incept`
        - :ref:`help_menu_context_incept_class_ex`
    """

    def __init__(self, container: ActionTriggerContainer) -> None:
        """
        Constructor

        Args:
            item (ActionTriggerItem): Menu item data.
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

    def _process_command(self, pop: ActionTriggerItem) -> None:
        """Process command"""
        if not pop.CommandURL:
            return
        pop.CommandURL = Shortcuts.get_url_script(pop.CommandURL)

    def _get_action_from_module(self, menu: dict, index: int) -> ActionTriggerItem | ActionTriggerSep | None:
        """
        Gets the Action Trigger item and looks up the command data from the command info.

        Args:
            menu (dict): Menu Data
            index (int): Current Index

        Raises:
            ValueError: If no text is found for the command

        Returns:
            ActionTriggerItem | ActionTriggerSep | None: Action Item if found, otherwise None

        Note:
            If the menu text is not found then the event ``action_item_module_no_text_found`` is raised.
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
                self.trigger_event("action_item_module_no_text_found", cargs)
                if cargs.cancel:
                    return None
                if "text" in cargs.event_data.menu:
                    cmd_text = cargs.event_data.menu["text"]

        if not cmd_text:
            raise ValueError(f"No text for command: {module_kind} {cmd}")

        menu_data = {
            "command": cmd,
        }
        menu_data.update(menu)
        # set the text back to the found command text
        menu_data["text"] = cmd_text
        return self._get_action_item(menu_data, index)

    def _get_action_item(self, menu: dict, index: int) -> ActionTriggerItem:
        """Get popup item data"""
        eargs = EventArgs(self)

        eargs.event_data = DotDict(**menu)
        eargs.event_data.index = index
        self.trigger_event("before_get_action_item", eargs)
        menu_data = cast(DotDict, eargs.event_data)
        text = menu_data.text
        command = cast(Union[str, "CommandDict"], menu_data.get("command", ""))
        help_command = menu_data.get("help_command", "")
        return ActionTriggerItem(command_url=Shortcuts.get_url_script(command), text=text, help_url=help_command)

    def get_action_item(self, menu: dict, index: int) -> ActionTriggerItem | ActionTriggerSep | None:
        """Get Action item"""
        if "module" in menu:
            item = self._get_action_from_module(menu, index)
        elif "text" in menu:
            text = menu["text"]
            if text == "-":
                item = ActionTriggerSep(int(menu.get("separator_type", 0)))
            else:
                item = self._get_action_item(menu, index)
        else:
            item = self._get_action_item(menu, index)
        return item

    def process(self, menu: dict, index: int) -> ActionTriggerItem | ActionTriggerSep | None:
        """Process menu item"""
        item = self.get_action_item(menu, index)
        if item is None:
            return None
        cargs = CancelEventArgs(source=self)
        cargs.event_data = DotDict(container=self._container, action_item=item)
        self.trigger_event("before_process", cargs)
        if cargs.cancel:
            return item

        self._container.insertByIndex(index, item)
        if isinstance(item, ActionTriggerSep):
            return item

        self._process_command(item)
        eargs = EventArgs.from_args(cargs)
        self.trigger_event("after_process", eargs)
        return item
