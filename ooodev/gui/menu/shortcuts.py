from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, Tuple

import uno
from ooo.dyn.awt.key_event import KeyEvent
from ooo.dyn.awt.key_modifier import KeyModifierEnum
from ooo.dyn.beans.property_concept import PropertyConceptEnum
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
    from ooodev.adapter.ui.ui_configuration_manager_comp import UIConfigurationManagerComp
    from com.sun.star.ui import XAcceleratorConfiguration


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

    def __init__(self, app: str = ""):
        self._app = app
        self._config = self._get_config()
        self._key_events = cast(Tuple[KeyEvent, ...], None)

    def _get_config(self) -> XAcceleratorConfiguration:
        if self._app:
            supp = TheModuleUIConfigurationManagerSupplierComp.from_lo()
            config = supp.get_ui_configuration_manager(self._app)
            return config.get_short_cut_manager().component
        else:
            return GlobalAcceleratorConfigurationComp.from_lo().component

    def __getitem__(self, index):
        return ShortCuts(index)

    def __contains__(self, item):
        cmd = self.get_by_shortcut(item)
        return bool(cmd)

    def __iter__(self):
        self._i = -1
        self._key_events = self._config.getAllKeyEvents()
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
    def get_url_script(cls, command: str | dict) -> str:
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
        cmd = self._config.getCommandByKeyEvent(key)
        shortcut = self._get_shortcut(key)
        return shortcut, cmd

    def get_all(self):
        """Get all events key"""
        events = [(self._get_info(k)) for k in self._config.getAllKeyEvents()]
        return events

    def get_by_command(self, command: str | dict):
        """Get shortcuts by command"""
        url = ShortCuts.get_url_script(command)
        key_events = self._config.getKeyEventsByCommand(url)
        shortcuts = [self._get_shortcut(k) for k in key_events]
        return shortcuts

    def get_by_shortcut(self, shortcut: str):
        """Get command by shortcut"""
        key_event = ShortCuts.to_key_event(shortcut)
        if key_event is None:
            error(f"Not exists shortcut: {shortcut}")
            return ""
        try:
            command = self._config.getCommandByKeyEvent(key_event)
        except NoSuchElementException:
            error(f"Not exists shortcut: {shortcut}")
            command = ""
        return command

    def set(self, shortcut: str, command: str | dict) -> bool:
        """Set shortcut to command

        :param shortcut: Shortcut like Shift+Ctrl+Alt+LETTER
        :type shortcut: str
        :param command: Command tu assign, 'UNOCOMMAND' or dict with macro info
        :type command: str or dict
        :return: True if set successfully
        :rtype: bool
        """
        result = True
        url = ShortCuts.get_url_script(command)
        key_event = ShortCuts.to_key_event(shortcut)
        if key_event is None:
            error(f"Not exists shortcut: {shortcut}")
            return False
        try:
            self._config.setKeyEvent(key_event, url)
            self._config.store()
        except Exception as e:
            error(e)
            result = False

        return result

    def remove_by_shortcut(self, shortcut: str):
        """Remove by shortcut"""
        key_event = ShortCuts.to_key_event(shortcut)
        key_event = ShortCuts.to_key_event(shortcut)
        if key_event is None:
            error(f"Not exists shortcut: {shortcut}")
            return False
        try:
            self._config.removeKeyEvent(key_event)
            result = True
        except NoSuchElementException:
            debug(f"No exists: {shortcut}")
            result = False
        return result

    def remove_by_command(self, command: str | dict):
        """Remove by shortcut"""
        url = ShortCuts.get_url_script(command)
        self._config.removeCommandFromAllKeyEvents(url)
        return

    def reset(self):
        """Reset configuration"""
        self._config.reset()
        self._config.store()
        return
