from __future__ import annotations
import copy
import json
from enum import Enum
from typing import Any, Dict, List, Callable, TYPE_CHECKING
import ooodev
from ooodev.gui.menu.popup.popup_item import PopupItem
from ooodev.gui.menu.popup_menu import PopupMenu
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.loader import lo as mLo
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.args.event_args import EventArgs
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.gui.menu.popup.popup_processor import PopupProcessor
from ooodev.io.json.json_encoder import JsonEncoder
from ooodev.utils.helper.dot_dict import DotDict


if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst


class PopupCreator(LoInstPropsPartial, EventsPartial, JsonEncoder):
    """
    Class for creating popup menu.

    This class can also be used to convert menu data to a format that can be used with JSON.

    Example:
        .. code-block:: python

            # ...
            menus = get_popup_menu()
            json_str = json.dumps(menus, cls=PopupCreator, indent=4)
            with open("popup_menu.json", "w") as f:
                f.write(json_str)

    See Also:
        - :ref:`help_popup_from_dict_or_json`
        - :ref:`help_popup_via_builder_item`
    """

    def __init__(self, lo_inst: LoInst | None = None, **kwargs) -> None:
        """
        Constructor

        Args:
            lo_inst (LoInst | None, optional): LibreOffice instance. Defaults to ``None``.
            kwargs (Any, optional): Additional keyword arguments. (Used by JsonEncoder)
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst)
        EventsPartial.__init__(self)
        JsonEncoder.__init__(self, **kwargs)
        self._lookups = {
            "text": "text",
            "command": "command",
            "style": "style",
            "checked": "checked",
            "enabled": "enabled",
            "default": "default",
            "help_command": "help_command",
            "help_text": "help_text",
            "tip_help_text": "tip_help_text",
            "shortcut": "shortcut",
            "submenu": "submenu",
        }

    def on_json_encode(self, obj: Any) -> Any:
        """
        Protected method to encode object to JSON.
        Can be overridden by subclasses.

        Args:
            obj (Any): Object to encode. This is the object that json is currently encoding.

        Returns:
            Any: The result of the encoding. The default is ``NULL_OBJ`` which means that the encoding is not handled.
        """
        if isinstance(obj, Enum):
            return int(obj)  # type: ignore
        return super().on_json_encode(obj)

    def _set_dict_from_popup_item(self, menu: dict[str, Any], pop: PopupItem) -> None:
        """Set dictionary from popup item"""
        if pop.is_separator():
            menu["text"] = "-"
            return
        if "text" in self.key_lookups:
            menu[self.key_lookups["text"]] = pop.text
        if "command" in self.key_lookups:
            menu[self.key_lookups["command"]] = pop.command
        if "style" in self.key_lookups:
            menu[self.key_lookups["style"]] = pop.style
        if "checked" in self.key_lookups:
            menu[self.key_lookups["checked"]] = pop.checked
        if "enabled" in self.key_lookups:
            menu[self.key_lookups["enabled"]] = pop.enabled
        if "default" in self.key_lookups:
            menu[self.key_lookups["default"]] = pop.default
        if "help_command" in self.key_lookups:
            menu[self.key_lookups["help_command"]] = pop.help_command
        if "help_text" in self.key_lookups:
            menu[self.key_lookups["help_text"]] = pop.help_text
        if "tip_help_text" in self.key_lookups:
            menu[self.key_lookups["tip_help_text"]] = pop.tip_help_text
        if "shortcut" in self.key_lookups:
            menu[self.key_lookups["shortcut"]] = pop.shortcut

    def _process_sub_menu(self, menus: list[dict[str, Any]]) -> None:
        """Insert submenu"""
        pm = PopupMenu.from_lo(lo_inst=self.lo_inst)
        # eargs = EventArgs(self)
        # eargs.event_data = {"popup_menu": parent_dict}
        # self.trigger_event("popup_created", eargs)

        for index, menu in enumerate(menus):
            submenu = menu.pop("submenu", False)
            mp = PopupProcessor(pm)
            mp.add_event_observers(self.event_observer)
            pop = mp.get_popup_item(menu, index)
            if pop is None:
                continue
            menu.clear()
            self._set_dict_from_popup_item(menu, pop)

            if submenu:
                self._process_sub_menu(submenu)
                menu[self.key_lookups["submenu"]] = submenu

    def _insert_sub_menu(self, parent: PopupMenu, parent_menu_id: int, menus: list[dict[str, Any]]) -> None:
        """Insert submenu"""
        pm = PopupMenu.from_lo(lo_inst=self.lo_inst)
        eargs = EventArgs(self)
        eargs.event_data = DotDict(popup_menu=pm)
        self.trigger_event("popup_created", eargs)

        for index, menu in enumerate(menus):
            submenu = menu.pop("submenu", False)
            mp = PopupProcessor(pm)
            mp.add_event_observers(self.event_observer)
            pop = mp.process(menu, index)
            if pop is None:
                continue
            if submenu:
                self._insert_sub_menu(pm, pop.menu_id, submenu)
        parent.set_popup_menu(parent_menu_id, pm)

    def create(self, menus: List[Dict[str, Any]]) -> PopupMenu:
        """
        Create popup menu.

        Args:
            menus (List[Dict[str, Any]]): Menu Data.
        """
        cpy = copy.deepcopy(menus)
        pm = PopupMenu.from_lo(lo_inst=self.lo_inst)
        eargs = EventArgs(self)
        eargs.event_data = DotDict(popup_menu=pm)
        self.trigger_event("popup_created", eargs)
        for index, menu in enumerate(cpy):
            submenu = menu.pop("submenu", False)
            mp = PopupProcessor(pm)
            mp.add_event_observers(self.event_observer)
            pop = mp.process(menu, index)
            if pop is None:
                continue
            if submenu:
                self._insert_sub_menu(pm, pop.menu_id, submenu)

        return pm

    # region subscribe

    def subscribe_popup_created(self, callback: Callable[[Any, EventArgs], None]) -> None:
        """
        Subscribe on popup created event.

        The callback ``event_data`` is a dictionary with keys:

        - ``popup_menu``: PopupMenu instance
        """
        self.subscribe_event("popup_created", callback)

    def unsubscribe_popup_created(self, callback: Callable[[Any, EventArgs], None]) -> None:
        """
        Unsubscribe on popup created event.
        """
        self.unsubscribe_event("popup_created", callback)

    def subscribe_before_process(self, callback: Callable[[Any, CancelEventArgs], None]) -> None:
        """
        Subscribe on before process event.

        The callback ``event_data`` is a dictionary with keys:

        - ``popup_menu``: PopupMenu instance
        - ``popup_item``: PopupItem instance
        """
        self.subscribe_event("before_process", callback)

    def subscribe_after_process(self, callback: Callable[[Any, EventArgs], None]) -> None:
        """
        Subscribe on after process event.

        The callback ``event_data`` is a dictionary with keys:

        - ``popup_menu``: PopupMenu instance
        - ``popup_item``: PopupItem instance
        """
        self.subscribe_event("after_process", callback)

    def unsubscribe_before_process(self, callback: Callable[[Any, CancelEventArgs], None]) -> None:
        """
        Unsubscribe on before process event.
        """
        self.unsubscribe_event("before_process", callback)

    def unsubscribe_after_process(self, callback: Callable[[Any, EventArgs], None]) -> None:
        """
        Unsubscribe on after process event.
        """
        self.unsubscribe_event("after_process", callback)

    def subscribe_module_no_text(self, callback: Callable[[Any, CancelEventArgs], None]) -> None:
        """
        Subscribe on no text found for module menu entry.

        This event occurs when a module menu entry is created using the ``module`` key and the text for the menu entry is not found.
        This event will not be raised if the ``module`` entry also provides a ``text`` key.
        A ``text`` key can be provided to the module entry to provide a valid menu text as a replacement if not found.

        The callback ``event_data`` is a dictionary with keys:

        - ``module_kind``: ModuleNamesKind
        - ``cmd``: Command as a string.
        - ``index``: Index as an integer.
        - ``menu``: Menu Data as a dictionary.

        The caller can set ``menu["text"]`` to provide a valid menu text.
        If the caller cancels the event then the menu item is not created.
        """
        self.subscribe_event("popup_module_no_text_found", callback)

    def unsubscribe_module_no_text(self, callback: Callable[[Any, CancelEventArgs], None]) -> None:
        """
        Unsubscribe on no text found for module menu entry.
        """
        self.unsubscribe_event("popup_module_no_text_found", callback)

    # endregion subscribe

    # region JSON

    def get_json_dict(self, menus: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Gets a dictionary that can be converted to JSON.
        This is an alternative to created json data. This is a more standard way that would not require this library to decode.

        The dictionary created by this method can also be used with the ``create`` method to create a popup menu.

        The menu dictionaries can have data such as ``{"command": ".uno:Cut", "module": ModuleNamesKind.SPREADSHEET_DOCUMENT}``.
        For json this sort of data does not work, this this method converts all menu data to a standard format that can be used with json.

        Args:
            menus (List[Dict[str, Any]]): Menu Data. This is can be the same data used to create a popup menu.

        Note:
            Even though menu data such as ``{"command": ".uno:Cut", "module": ModuleNamesKind.SPREADSHEET_DOCUMENT}`` is not valid for json.
            It can still be encoded using this ``PopupCreator`` class.

            Example:
                .. code-block:: python

                    # ...
                    menus = get_popup_menu()
                    json_str = json.dumps(menus, cls=PopupCreator, indent=4)
                    with open("popup_menu.json", "w") as f:
                        f.write(json_str)
        """
        result = []
        pm = PopupMenu.from_lo(lo_inst=self.lo_inst)
        for index, menu in enumerate(menus):
            cpy = copy.deepcopy(menu)
            submenu = cpy.pop("submenu", False)

            mp = PopupProcessor(pm)
            mp.add_event_observers(self.event_observer)
            pop = mp.get_popup_item(cpy, index)
            if pop is None:
                continue
            cpy.clear()
            self._set_dict_from_popup_item(cpy, pop)
            if submenu:
                self._process_sub_menu(submenu)
                cpy[self.key_lookups["submenu"]] = submenu
            result.append(cpy)
        return result

    def json_dumps(self, menus: List[Dict[str, Any]], dynamic: bool = False) -> str:
        """
        Get JSON data.

        Args:
            menus (List[Dict[str, Any]]): Menu Data.

        Returns:
            str: JSON data.
        """
        version = ooodev.get_version()
        if dynamic:
            menu_data = menus
        else:
            menu_data = self.get_json_dict(menus)
        data = {"id": "ooodev.popup_menu", "version": version, "dynamic": dynamic, "menus": menu_data}
        if dynamic:
            return json.dumps(data, cls=PopupCreator, indent=4)
        return json.dumps(data, indent=4)

    def json_dump(self, file: Any, menus: List[Dict[str, Any]], dynamic: bool = False) -> None:
        """
        Dump JSON data to file.

        Args:
            file (Any): File path.
            menus (List[Dict[str, Any]]): Menu Data.
            dynamic (bool, optional): Dynamic data. Defaults to ``False``.
        """
        with open(file, "w") as f:
            f.write(self.json_dumps(menus, dynamic=dynamic))

    @staticmethod
    def json_loads(json_str: str, **kwargs) -> List[Dict[str, Any]]:
        """
        Load JSON data.

        Args:
            json_str (str): JSON data.

        Returns:
            List[Dict[str, Any]]: Menu Data.
        """
        data = json.loads(json_str, **kwargs)
        if data["id"] != "ooodev.popup_menu":
            raise ValueError("Invalid JSON data")
        return data["menus"]

    @staticmethod
    def json_load(json_file: Any, **kwargs) -> List[Dict[str, Any]]:
        """
        Load JSON data from file.

        Args:
            json_file (str): JSON file path.

        Returns:
            List[Dict[str, Any]]: Menu Data.
        """
        with open(json_file, "r") as f:
            data = json.load(f, **kwargs)
        if data["id"] != "ooodev.popup_menu":
            raise ValueError("Invalid JSON data")
        return data["menus"]

    # endregion JSON

    # region Properties
    @property
    def key_lookups(self) -> Dict[str, str]:
        """
        Get key lookups.

        Returns:
            Dict[str, Any]: Key lookups.
        """
        return self._lookups

    @key_lookups.setter
    def key_lookups(self, value: Dict[str, str]) -> None:
        """
        Set key lookups.

        Args:
            value (Dict[str, str]): Key lookups.
        """
        self._lookups = value

    # endregion Properties
