from __future__ import annotations
from typing import cast, Dict, List, TYPE_CHECKING, Tuple, Union

import uno
from ooo.dyn.awt.key_event import KeyEvent
from ooo.dyn.awt.key_modifier import KeyModifierEnum
from com.sun.star.awt import Key
from com.sun.star.container import NoSuchElementException

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

if TYPE_CHECKING:
    from ooodev.adapter.ui.accelerator_configuration_comp import AcceleratorConfigurationComp
    from ooodev.loader.inst.lo_inst import LoInst


class Shortcuts(LoInstPropsPartial):
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
        self._command_dict = cast(Dict[str, List[str]], None)
        self._logger = NamedLogger(name="ShortCuts")

    def _get_config(self) -> Union[AcceleratorConfigurationComp, GlobalAcceleratorConfigurationComp]:
        if self._app:
            supp = TheModuleUIConfigurationManagerSupplierComp.from_lo(lo_inst=self.lo_inst)
            config = supp.get_ui_configuration_manager(self._app)
            return config.get_short_cut_manager()
        else:
            return GlobalAcceleratorConfigurationComp.from_lo(lo_inst=self.lo_inst)

    def __getitem__(self, app: str | Service):
        return Shortcuts(app)

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
                if not m:
                    # could be empty string
                    continue
                key_event.Modifiers += cls.MODIFIERS[m.lower()]
            key_event.KeyCode = getattr(Key, keys[-1].upper())
        except Exception as e:
            logger.error("Exception occured", exc_info=True)
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

    def _get_shortcut(self, key: KeyEvent):
        """Get shortcut for key event"""
        # print(k.KeyCode, str(k.KeyChar), k.KeyFunc, k.Modifiers)
        shortcut = f"{self.COMBINATIONS[key.Modifiers]}+{self.KEYS[key.KeyCode]}"
        return shortcut

    def _get_info(self, key: KeyEvent) -> Tuple[str, str]:
        """Get shortcut and command"""
        cmd = self._config.get_command_by_key_event(key)
        shortcut = self._get_shortcut(key)
        return shortcut, cmd

    def get_all(self) -> List[Tuple[str, str]]:
        """
        Get all events key.

        Returns:
            List[Tuple[str, str]]: List of tuples with shortcut and command.
        """
        events = [(self._get_info(k)) for k in self._config.get_all_key_events()]
        return events

    def _get_by_command_dict(self, url: str) -> List[str]:
        # for unknown reason LibreOffice does not return the command for most urls.
        # This method is a workaround to get the command by url.
        if self._command_dict is None:
            self._command_dict: Dict[str, List[str]] = {}
            for key in self._config.get_all_key_events():
                try:
                    cmd = self._config.get_command_by_key_event(key)
                except Exception:
                    continue
                if cmd in self._command_dict:
                    self._command_dict[cmd].append(self._get_shortcut(key))
                else:
                    self._command_dict[cmd] = [self._get_shortcut(key)]
        if url in self._command_dict:
            return self._command_dict[url]
        return []

    def get_by_command(self, command: str | Dict[str, str]) -> List[str]:
        """
        Get shortcuts by command.

        Args:
            command (str | dict): Command to search, 'UNOCOMMAND' or dict with macro info.

        Returns:
            List[str]: List of shortcuts or empty list of not found.
        """
        url = Shortcuts.get_url_script(command)
        try:
            key_events = self._config.get_key_events_by_command(url)
            shortcuts = [self._get_shortcut(k) for k in key_events]
        except NoSuchElementException:
            # fallback on workaround
            shortcuts = self._get_by_command_dict(url)
        return shortcuts

    def get_by_shortcut(self, shortcut: str) -> str:
        """Get command by shortcut"""
        key_event = Shortcuts.to_key_event(shortcut)
        if key_event is None:
            self._logger.warning(f"get_by_shortcut() - Not exists shortcut: {shortcut}")
            return ""
        try:
            command = self._config.get_command_by_key_event(key_event)
        except NoSuchElementException:
            self._logger.warning(f"Not exists shortcut: {shortcut}")
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
        url = Shortcuts.get_url_script(command)
        key_event = Shortcuts.to_key_event(shortcut)
        if key_event is None:
            self._logger.warning(f"Not exists shortcut: {shortcut}")
            return False
        try:
            self._config.set_key_event(key_event, url)
            self._config.store()
        except Exception as e:
            self._logger.error(e)
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
        key_event = Shortcuts.to_key_event(shortcut)
        if key_event is None:
            self._logger.warning(f"Not exists shortcut: {shortcut}")
            return False
        try:
            self._config.remove_key_event(key_event)
            result = True
        except NoSuchElementException:
            self._logger.debug(f"No exists: {shortcut}")
            result = False
        return result

    def remove_by_command(self, command: str | Dict[str, str]):
        """
        Remove by shortcut.

        Args:
            command (str | dict): Command to remove, 'UNOCOMMAND' or dict with macro info
        """
        url = Shortcuts.get_url_script(command)
        self._config.remove_command_from_all_key_events(url)
        return

    def reset(self):
        """Reset configuration"""
        self._config.reset()  # type: ignore
        self._config.store()
        return
