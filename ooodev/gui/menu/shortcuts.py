from __future__ import annotations
from typing import cast, Dict, TYPE_CHECKING, Tuple, Union

import uno
from ooo.dyn.awt.key_event import KeyEvent
from ooo.dyn.awt.key_modifier import KeyModifierEnum
from com.sun.star.awt import Key
from com.sun.star.container import NoSuchElementException

from ooodev.loader.inst.service import Service
from ooodev.adapter.ui.global_accelerator_configuration_comp import GlobalAcceleratorConfigurationComp
from ooodev.adapter.ui.the_module_ui_configuration_manager_supplier_comp import (
    TheModuleUIConfigurationManagerSupplierComp,
)
from ooodev.macro.script.macro_script import MacroScript
from ooodev.io.logging import error, debug

if TYPE_CHECKING:
    from ooodev.adapter.ui.accelerator_configuration_comp import AcceleratorConfigurationComp


class ShortCuts:
    """Class for manager shortcuts"""

    KEYS = {getattr(Key, k): k for k in dir(Key)}
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

    def __init__(self, app: str | Service = ""):
        """
        Constructor

        Args:
            app (str | Service, optional): App Name such as ``Service.CALC``. Defaults to "".
                If no app is provided, the global shortcuts will be used.

        Hint:
            - ``Service`` is an enum and can be imported from ``ooodev.loader.inst.service``
        """
        self._app = str(app)
        self._config = self._get_config()
        self._key_events = cast(Tuple[KeyEvent, ...], None)

    def _get_config(self) -> Union[AcceleratorConfigurationComp, GlobalAcceleratorConfigurationComp]:
        if self._app:
            supp = TheModuleUIConfigurationManagerSupplierComp.from_lo()
            config = supp.get_ui_configuration_manager(self._app)
            return config.get_short_cut_manager()
        else:
            return GlobalAcceleratorConfigurationComp.from_lo()

    def __getitem__(self, app: str | Service):
        return ShortCuts(app)

    def __contains__(self, item):
        cmd = self.get_by_shortcut(item)
        return bool(cmd)

    def __iter__(self):
        self._i = -1
        self._key_events = self._config.get_all_key_events()
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
    def to_key_event(cls, shortcut: str):
        """Convert from string shortcut (Shift+Ctrl+Alt+LETTER) to KeyEvent"""
        key_event = KeyEvent()
        keys = shortcut.split("+")
        try:
            for m in keys[:-1]:
                key_event.Modifiers += cls.MODIFIERS[m.lower()]
            key_event.KeyCode = getattr(Key, keys[-1].upper())
        except Exception as e:
            error(e)
            key_event = None
        return key_event

    @classmethod
    def get_url_script(cls, command: str | Dict[str, str]) -> str:
        """Get uno command or url for macro"""
        url = command
        if isinstance(url, str) and not url.startswith(".uno:"):
            url = f".uno:{command}"
        elif isinstance(url, dict):
            url = MacroScript.get_url_script(**url)
        return url

    def _get_shortcut(self, k):
        """Get shortcut for key event"""
        # ~ print(k.KeyCode, str(k.KeyChar), k.KeyFunc, k.Modifiers)
        shortcut = f"{self.COMBINATIONS[k.Modifiers]}+{self.KEYS[k.KeyCode]}"
        return shortcut

    def _get_info(self, key):
        """Get shortcut and command"""
        cmd = self._config.get_command_by_key_event(key)
        shortcut = self._get_shortcut(key)
        return shortcut, cmd

    def get_all(self):
        """Get all events key"""
        events = [(self._get_info(k)) for k in self._config.get_all_key_events()]
        return events

    def get_by_command(self, command: str | Dict[str, str]):
        """Get shortcuts by command"""
        url = ShortCuts.get_url_script(command)
        key_events = self._config.get_key_events_by_command(url)
        shortcuts = [self._get_shortcut(k) for k in key_events]
        return shortcuts

    def get_by_shortcut(self, shortcut: str):
        """Get command by shortcut"""
        key_event = ShortCuts.to_key_event(shortcut)
        if key_event is None:
            error(f"Not exists shortcut: {shortcut}")
            return ""
        try:
            command = self._config.get_command_by_key_event(key_event)
        except NoSuchElementException:
            error(f"Not exists shortcut: {shortcut}")
            command = ""
        return command

    def set(self, shortcut: str, command: str | Dict[str, str]) -> bool:
        """
        Set shortcut to command

        Args:
            shortcut (str): Shortcut like Shift+Ctrl+Alt+LETTER
            command (str | dict): Command to assign, 'UNOCOMMAND' or dict with macro info

        Returns:
            bool: True if set successfully
        """
        result = True
        url = ShortCuts.get_url_script(command)
        key_event = ShortCuts.to_key_event(shortcut)
        if key_event is None:
            error(f"Not exists shortcut: {shortcut}")
            return False
        try:
            self._config.set_key_event(key_event, url)
            self._config.store()
        except Exception as e:
            error(e)
            result = False

        return result

    def remove_by_shortcut(self, shortcut: str) -> bool:
        """
        Remove by shortcut

        Args:
            shortcut (str): Shortcut like Shift+Ctrl+Alt+LETTER

        Returns:
            bool: ``True`` if removed successfully
        """
        key_event = ShortCuts.to_key_event(shortcut)
        if key_event is None:
            error(f"Not exists shortcut: {shortcut}")
            return False
        try:
            self._config.remove_key_event(key_event)
            result = True
        except NoSuchElementException:
            debug(f"No exists: {shortcut}")
            result = False
        return result

    def remove_by_command(self, command: str | Dict[str, str]):
        """
        Remove by shortcut.

        Args:
            command (str | dict): Command to remove, 'UNOCOMMAND' or dict with macro info
        """
        url = ShortCuts.get_url_script(command)
        self._config.remove_command_from_all_key_events(url)
        return

    def reset(self):
        """Reset configuration"""
        self._config.reset()
        self._config.store()
        return
