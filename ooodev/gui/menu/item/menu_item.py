from __future__ import annotations
from typing import Tuple, List, TYPE_CHECKING
import contextlib
import uno
from ooodev.loader.inst.service import Service
from ooodev.loader import lo as mLo
from ooodev.utils.kind.item_style_kind import ItemStyleKind
from com.sun.star.beans import PropertyValue
from ooodev.gui.menu.item.menu_item_base import MenuItemBase
from ooodev.adapter.container.index_access_comp import IndexAccessComp
from ooodev.macro.script.macro_script import MacroScript
from ooodev.io.log.named_logger import NamedLogger

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst


class MenuItem(MenuItemBase):
    def __init__(
        self,
        *,
        data: Tuple[Tuple[PropertyValue, ...], ...],
        owner: IndexAccessComp,
        app: str | Service = "",
        lo_inst: LoInst | None = None,
    ):
        """
        Constructor

        Args:
            component (XIndexAccess): UNO Object containing menu item properties.
        """
        super().__init__(data=data, owner=owner, app=app, lo_inst=lo_inst)
        if self.command:
            lg_name = f"{self.__class__.__name__} ({self.command})"
        else:
            lg_name = self.__class__.__name__
        self.__logger = NamedLogger(lg_name)

    def execute(self) -> bool:
        """Execute menu item"""
        cmd = self.command
        if not cmd:
            return False
        supported_prefixes = tuple(self.lo_inst.get_supported_dispatch_prefixes())
        if cmd.startswith(supported_prefixes):
            self.lo_inst.dispatch_cmd(cmd)
            return True
        try:
            script = MacroScript.get_script(cmd)
            script.invoke((), None, None)  # type: ignore
        except Exception as e:
            self.__logger.error(f"Error executing menu item with command value of '{cmd}': {e}")
            return False
        return False

    def __str__(self) -> str:
        return self.command

    def __repr__(self) -> str:
        if self.label:
            return f'<MenuItem(command="{self.command}", label="{self.label}")>'
        return f'<MenuItem(command="{self.command}")>'

    def get_shortcuts(self) -> List[str]:
        """Get shortcuts"""
        from ooodev.gui.menu.shortcuts import Shortcuts

        sc = Shortcuts(app=self._app, lo_inst=self.lo_inst)
        return sc.get_by_command(self.command)

    @property
    def command(self) -> str:
        """Get/Set command"""
        return self.data_dict.get("CommandURL", "")

    @command.setter
    def command(self, value: str):
        """Set command"""
        self.data_dict["CommandURL"] = value

    @property
    def label(self) -> str:
        """Get/Set label"""
        return self.data_dict.get("Label", "")

    @label.setter
    def label(self, value: str):
        """Set label"""
        self.data_dict["Label"] = value

    @property
    def help_url(self) -> str:
        """Get/Set help text"""
        return self.data_dict.get("HelpURL", "")

    @help_url.setter
    def help_url(self, value: str):
        """Set help text"""
        self.data_dict["HelpURL"] = value

    @property
    def style(self) -> ItemStyleKind:
        """
        Get/Set style

        Hint:
            - ``ItemStyleKind`` is an enum and can be imported from ``ooodev.utils.kind.item_style_kind``.
        """
        return ItemStyleKind(int(self._menu_data.get("Style", 0)))

    @style.setter
    def style(self, value: int | ItemStyleKind):
        """Set style"""
        self._menu_data["Style"] = int(value)

    # @property
    # def sub_menu(self) -> IndexAccessComp | None:
    #     """Get/Set help text"""
    #     obj = self._menu_data.get("ItemDescriptorContainer", "")
    #     if obj is None:
    #         return None
    #     menu = Menu(self._config, self._menus, self._app, obj)
