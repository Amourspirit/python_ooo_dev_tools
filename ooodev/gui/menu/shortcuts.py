from __future__ import annotations
from typing import cast, Dict, List, TYPE_CHECKING, Tuple, Union

import uno
from ooo.dyn.awt.key_event import KeyEvent
from ooo.dyn.awt.key_modifier import KeyModifierEnum
from com.sun.star.awt import Key
from com.sun.star.container import NoSuchElementException

from ooodev.mock import mock_g
from ooodev.adapter.ui.global_accelerator_configuration_comp import GlobalAcceleratorConfigurationComp
from ooodev.adapter.ui.the_module_ui_configuration_manager_supplier_comp import (
    TheModuleUIConfigurationManagerSupplierComp,
)
from ooodev.io.log import logging as logger
from ooodev.io.log.named_logger import NamedLogger
from ooodev.loader import lo as mLo
from ooodev.loader.inst.service import Service
from ooodev.macro.script.macro_script import MacroScript
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.cache.lru_cache import LRUCache
from ooodev.utils.string.str_list import StrList

if TYPE_CHECKING:
    from ooodev.adapter.ui.accelerator_configuration_comp import AcceleratorConfigurationComp
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.gui.menu.common.command_dict import CommandDict


class Shortcuts(LoInstPropsPartial):
    """
    Class for manager shortcuts.

    See Also:

        - :ref:`help_working_with_shortcuts`
    """

    if mock_g.DOCS_BUILDING:
        # When docs are building it may not have access to uno.
        # Key is a uno object and it is not available so this will cause errors when building docs.
        KEYS: Dict[int, str] = {}
        """
        Keys dictionary. This is a dictionary that is built at runtime with the keys and values of the ``Key`` class. The dictionary Keys are the integer values of the keys and the values are the key names.
        
        See `API Key <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1awt_1_1Key.html>`_
        """
    else:
        KEYS: Dict[int, str] = {getattr(Key, k): k for k in dir(Key)}
        """The dictionary Keys are the integer values of the keys and the values are the key names."""

    MODIFIERS = {
        "shift": KeyModifierEnum.SHIFT.value,
        "ctrl": KeyModifierEnum.MOD1.value,
        "alt": KeyModifierEnum.MOD2.value,
        "ctrlmac": KeyModifierEnum.MOD3.value,
    }
    COMBINATIONS = {
        0: "",
        1: "shift",
        2: "ctrl",
        4: "alt",
        8: "ctrlmac",
        3: "shift+ctrl",
        5: "shift+alt",
        9: "shift+ctrlmac",
        6: "ctrl+alt",
        10: "ctrl+ctrlmac",
        12: "alt+ctrlmac",
        7: "shift+ctrl+alt",
        11: "shift+ctrl+ctrlmac",
        13: "shift+alt+ctrlmac",
        14: "ctrl+alt+ctrlmac",
        15: "shift+ctrl+alt+ctrlmac",
    }

    def __init__(self, app: str | Service = "", lo_inst: LoInst | None = None):
        """
        Constructor

        Args:
            app (str | Service, optional): App Name such as ``Service.CALC``. Defaults to "".
                If no app is provided, the global shortcuts will be used.

        Hint:
            - ``Service`` is an enum and can be imported from ``ooodev.loader.inst.service``
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst)
        self._app = str(app)
        self._config = self._get_config()
        self._key_events = cast(Tuple[KeyEvent, ...], None)
        self._logger = NamedLogger(name="Shortcuts")
        self._cache = LRUCache(50)

    def _get_config(self) -> Union[AcceleratorConfigurationComp, GlobalAcceleratorConfigurationComp]:
        if self._app:
            key = f"Shortcuts_get_ui_configuration_manager_{self._app}"
            if key in self.lo_inst.cache:
                return self.lo_inst.cache[key]
            supp = TheModuleUIConfigurationManagerSupplierComp.from_lo(lo_inst=self.lo_inst)

            config = supp.get_ui_configuration_manager(self._app)
            self.lo_inst.cache[key] = config.get_short_cut_manager()
            return cast("AcceleratorConfigurationComp", self.lo_inst.cache[key])
        else:
            return GlobalAcceleratorConfigurationComp.from_lo(lo_inst=self.lo_inst)

    def _get_all_key_events(self) -> Tuple[KeyEvent, ...]:
        key = "_get_all_key_events"
        if key in self._cache:
            return self._cache[key]
        key_events = self._config.get_all_key_events()
        self._cache[key] = key_events
        return key_events

    def __getitem__(self, app: str | Service):
        return Shortcuts(app)

    def __contains__(self, item):
        cmd = self.get_by_shortcut(item)
        return bool(cmd)

    def __iter__(self):
        self._i = -1
        self._key_events = self._get_all_key_events()
        return self

    def __next__(self):
        if self._key_events is None:
            raise StopIteration
        self._i += 1
        try:
            event = self._key_events[self._i]
            event = self._get_info(event)
        except IndexError:
            self._key_events = cast(Tuple[KeyEvent, ...], None)
            raise StopIteration

        return event

    @classmethod
    def to_key_event(cls, shortcut: str) -> KeyEvent | None:
        """Convert from string shortcut (Shift+Ctrl+Alt+LETTER) to KeyEvent"""
        key_event = KeyEvent()
        keys = shortcut.split("+")
        try:
            for m in keys[:-1]:
                if not m:
                    # could be empty string
                    continue
                key_event.Modifiers += cls.MODIFIERS[m.lower()]
            key_event.KeyCode = getattr(Key, keys[-1].upper())
        except Exception:
            logger.error("Exception occured", exc_info=True)
            key_event = None
        return key_event

    @classmethod
    def from_key_event(cls, key_event: KeyEvent) -> str:
        """Convert from KeyEvent to string shortcut"""
        shortcut = ""
        for m in cls.MODIFIERS:
            if key_event.Modifiers & cls.MODIFIERS[m]:
                shortcut += f"{m.capitalize()}+"
        shortcut += cls.KEYS[key_event.KeyCode]
        return shortcut

    @classmethod
    def get_url_script(cls, command: Union[str, CommandDict]) -> str:
        """
        Get uno command or url for macro.

        Args:
            command (str | CommandDict): Command to search, 'UNOCOMMAND' or dict with macro info.

        Returns:
            str: Url for macro or uno command or custom command.

        Note:
            If ``command`` is passed in a a string and it starts with ``.custom:``
            then it will be returned with the ``.custom:`` prefix dropped.

            The ``.custom:`` prefix is used to indicate that the command is a custom command
            and can be used in a menu callback to capture user clicks.
        """
        url = command
        if isinstance(url, str) and not url.startswith(".uno:"):
            if url.startswith(".custom:"):
                url = url[8:]
            else:
                url = f".uno:{command}"
        elif isinstance(url, dict):
            url = MacroScript.get_url_script(**url)
        return url

    def get_shortcut(self, key: KeyEvent) -> str:
        """
        Get shortcut for key event.

        Args:
            key (KeyEvent): Key event

        Returns:
            str: Shortcut like Shift+Ctrl+Alt+LETTER
        """
        # print(k.KeyCode, str(k.KeyChar), k.KeyFunc, k.Modifiers)
        shortcut = f"{self.COMBINATIONS[key.Modifiers]}+{self.KEYS[key.KeyCode]}"
        return shortcut

    def _get_info(self, key: KeyEvent) -> Tuple[str, str]:
        """Get shortcut and command"""
        cmd = self._config.get_command_by_key_event(key)
        shortcut = self.get_shortcut(key)
        return shortcut, cmd

    def get_all(self) -> List[Tuple[str, str]]:
        """
        Get all events key.

        Returns:
            List[Tuple[str, str]]: List of tuples with shortcut and command.
        """
        events = [(self._get_info(k)) for k in self._get_all_key_events()]
        return events

    def _get_by_command_dict(self, url: str) -> StrList:
        # for unknown reason LibreOffice does not return the command for most urls.
        # This method is a workaround to get the command by url.
        key = "_get_by_command_dict"
        if key in self._cache:
            command_dict = cast(Dict[str, List[str]], self._cache[key])
        else:
            command_dict: Dict[str, List[str]] = {}
            for key in self._get_all_key_events():
                try:
                    cmd = self._config.get_command_by_key_event(key)
                except Exception:
                    continue
                if cmd in command_dict:
                    command_dict[cmd].append(self.get_shortcut(key))
                else:
                    command_dict[cmd] = [self.get_shortcut(key)]
            self._cache[key] = command_dict
        if url in command_dict:
            return StrList(command_dict[url])
        return StrList()

    def get_by_command(self, command: Union[str, CommandDict]) -> StrList:
        """
        Get shortcuts by command.

        Args:
            command (str | dict): Command to search, 'UNOCOMMAND' or dict with macro info.

        Returns:
            List[str]: List of shortcuts or empty list of not found.
        """
        url = Shortcuts.get_url_script(command)
        key = f"Shortcuts_get_by_command_{url}"
        if key in self._cache:
            return self._cache[key]
        try:
            key_events = self._config.get_key_events_by_command(url)
            shortcuts = StrList([self.get_shortcut(k) for k in key_events])
        except NoSuchElementException:
            # fallback on workaround
            shortcuts = self._get_by_command_dict(url)
        self._cache[key] = shortcuts
        return shortcuts

    def get_by_shortcut(self, shortcut: str) -> str:
        """Get command by shortcut"""
        key = f"get_by_shortcut_{shortcut}"
        if key in self._cache:
            return self._cache[key]
        key_event = Shortcuts.to_key_event(shortcut)
        if key_event is None:
            self._logger.warning(f"get_by_shortcut() - Not exists shortcut: {shortcut}")
            return ""
        try:
            command = self._config.get_command_by_key_event(key_event)
        except NoSuchElementException:
            self._logger.warning(f"Not exists shortcut: {shortcut}")
            command = ""
        if command:
            self._cache[key] = command
        return command

    def get_by_key_event(self, key_event: KeyEvent) -> str:
        """Get command by key event"""
        sc = self.get_shortcut(key_event)
        return self.get_by_shortcut(sc)

    def set(self, shortcut: str, command: Union[str, CommandDict], save: bool = True) -> bool:
        """
        Set shortcut to command

        Args:
            shortcut (str): Shortcut like Shift+Ctrl+Alt+LETTER
            command (str | CommandDict): Command to assign, 'UNOCOMMAND' or dict with macro info.
            save (bool, optional): Save configuration causing it to persist. Defaults to ``True``.

        Returns:
            bool: True if set successfully
        """
        result = True
        url = Shortcuts.get_url_script(command)
        key_event = Shortcuts.to_key_event(shortcut)
        if key_event is None:
            self._logger.warning(f"Not exists shortcut: {shortcut}")
            return False
        try:
            self._config.set_key_event(key_event, url)
            if save:
                self._config.store()
        except Exception as e:
            self._logger.error(e)
            result = False

        return result

    def remove_by_shortcut(self, shortcut: str, save: bool = False) -> bool:
        """
        Remove by shortcut

        Args:
            shortcut (str): Shortcut like Shift+Ctrl+Alt+LETTER
            save (bool, optional): Save configuration causing it to persist. Defaults to ``False``.

        Returns:
            bool: ``True`` if removed successfully
        """
        key_event = Shortcuts.to_key_event(shortcut)
        if key_event is None:
            self._logger.warning(f"Not exists shortcut: {shortcut}")
            return False
        try:
            self._config.remove_key_event(key_event)
            if save:
                self._config.store()
            result = True
        except NoSuchElementException:
            self._logger.debug(f"No exists: {shortcut}")
            result = False
        if result:
            self._cache.clear()
        return result

    def remove_by_command(self, command: Union[str, CommandDict], save: bool = False):
        """
        Remove by shortcut.

        Args:
            command (str | CommandDict): Command to remove, 'UNOCOMMAND' or dict with macro info.
            save (bool, optional): Save configuration causing it to persist. Defaults to ``False``.
        """
        url = Shortcuts.get_url_script(command)
        self._config.remove_command_from_all_key_events(url)
        if save:
            self._config.store()
        self._cache.clear()
        return

    def reset(self) -> None:
        """Reset configuration"""
        self._config.reset()  # type: ignore
        self._config.store()
        return
