from __future__ import annotations
from typing import Any, Tuple, List, TYPE_CHECKING
import contextlib
import uno
from com.sun.star.beans import PropertyValue
from ooo.dyn.util.url import URL

from ooodev.adapter.container.index_access_comp import IndexAccessComp
from ooodev.gui.menu.comp.dispatch_comp import DispatchComp
from ooodev.gui.menu.item.menu_item_base import MenuItemBase
from ooodev.gui.menu.item.menu_item_kind import MenuItemKind
from ooodev.io.log.named_logger import NamedLogger
from ooodev.loader.inst.service import Service
from ooodev.macro.script.macro_script import MacroScript
from ooodev.utils.kind.item_style_kind import ItemStyleKind
from ooodev.utils.string.str_list import StrList

if TYPE_CHECKING:
    from ooodev.gui.menu.menu import Menu
    from ooodev.loader.inst.lo_inst import LoInst


class MenuItem(MenuItemBase):
    """Menu Item"""

    def __init__(
        self,
        *,
        menu: Menu,
        data: Tuple[Tuple[PropertyValue, ...], ...],
        owner: IndexAccessComp,
        app: str | Service = "",
        lo_inst: LoInst | None = None,
    ):
        """
        Constructor

        Args:
            data (Tuple[Tuple[PropertyValue, ...], ...]): UNO Object containing menu item properties.
            owner (IndexAccessComp): Parent menu.
            app (str | Service, optional): Name LibreOffice module. Defaults to "".
            lo_inst (LoInst | None, optional): Lo Instance. Defaults to Current Lo Instance.
        """
        super().__init__(data=data, menu=menu, owner=owner, app=app, lo_inst=lo_inst)
        if self.command:
            lg_name = f"{self.__class__.__name__} ({self.command})"
        else:
            lg_name = self.__class__.__name__
        self.__logger = NamedLogger(lg_name)

    def execute(self, in_thread: bool = False) -> bool:
        """Execute menu item"""
        cmd = self.command
        if not cmd:
            return False
        supported_prefixes = tuple(self.lo_inst.get_supported_dispatch_prefixes())
        if cmd.startswith(supported_prefixes):
            self.lo_inst.dispatch_cmd(cmd, in_thread=in_thread)
            return True
        try:
            _ = MacroScript.call_url(cmd, in_thread=in_thread)
            script = MacroScript.get_script(cmd)
            script.invoke((), None, None)  # type: ignore
            return True
        except Exception as e:
            self.__logger.error(f"Error executing menu item with command value of '{cmd}': {e}")
        return False

    def __str__(self) -> str:
        return self.command

    def __repr__(self) -> str:
        if self.label:
            return f'<{self.__class__.__name__}(command="{self.command}", label="{self.label}, kind={str(self.item_kind)}")>'
        return f'<{self.__class__.__name__}(command="{self.command}", kind={str(self.item_kind)})>'

    def get_shortcuts(self) -> StrList:
        """Get shortcuts"""
        from ooodev.gui.menu.shortcuts import Shortcuts

        sc = Shortcuts(app=self._app, lo_inst=self.lo_inst)
        return sc.get_by_command(self.command)

    def _get_dispatch(self) -> Any:
        from ooodev.adapter.util.url_transformer_comp import URLTransformerComp

        with contextlib.suppress(Exception):
            self._dispatch = None
            frame = self.lo_inst.get_frame()
            if frame is None:
                return None

            ts = URLTransformerComp.from_lo(lo_inst=self.lo_inst)
            url = URL(Complete=self.command)
            _, url = ts.parse_smart(url, ".uno:")
            return frame.queryDispatch(url, "", 0)  # type: ignore
        return None

    def get_dispatch(self) -> DispatchComp | None:
        """
        Get dispatch object.

        If the menu item is disabled, this method will return None.
        """
        dispatch = self._get_dispatch()
        if dispatch is None:
            return None
        return DispatchComp(dispatch)

    def is_enabled(self) -> bool:
        """Check if menu item is enabled"""
        return self._get_dispatch() is not None

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

    @property
    def item_kind(self) -> MenuItemKind:
        """
        Get item kind.

        Returns:
            MenuItemKind: ``MenuItemKind.ITEM``.

        Hint:
            - ``MenuItemKind`` is an enum and can be imported from ``ooodev.gui.menu.item.menu_item_kind``.
        """
        return MenuItemKind.ITEM
