from __future__ import annotations
from re import T
from typing import Any, cast, TYPE_CHECKING, Callable

import contextlib
import uno
from com.sun.star.awt import XPopupMenu
from ooo.dyn.awt.menu_item_type import MenuItemType

from ooodev.mock import mock_g
from ooodev.events.args.event_args import EventArgs
from ooodev.events.args.generic_args import GenericArgs
from ooodev.adapter.component_prop import ComponentProp
from ooodev.adapter.awt.popup_menu_partial import PopupMenuPartial
from ooodev.adapter.awt.menu_events import MenuEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs

if TYPE_CHECKING:
    from com.sun.star.awt import PopupMenu
    from ooodev.loader.inst.lo_inst import LoInst
    from typing_extensions import Self


class PopupMenuComp(ComponentProp, PopupMenuPartial, MenuEvents):
    """
    Class for managing PopupMenu Component.

    Events for this component are managed by the ``MenuEvents`` class.

    Event Callbacks are passed the following positional arguments:
        - src: The source of the event.
        - event: The event arguments. The event_data attribute contains ``com.sun.star.awt.MenuEvent``.
        - menu: The popup menu component that triggered the event.

    Example:
        .. code-block:: python

            def on_menu_Highlighted(src: Any, event: EventArgs,  menu: PopupMenuComp) -> None:
                print("Menu Highlighted")
                me = cast("MenuEvent", event.event_data)
                print("MenuId", me.MenuId)
    """

    # pylint: disable=unused-argument

    # region Dunder Methods
    def __init__(self, component: XPopupMenu) -> None:
        """
        Constructor

        Args:
            component (UnoControlDialog): UNO Component that supports `com.sun.star.awt.PopupMenu`` service.
        """
        # pylint: disable=no-member
        ComponentProp.__init__(self, component)
        PopupMenuPartial.__init__(self, component=component)
        generic_args = GenericArgs(menu=self)
        MenuEvents.__init__(self, trigger_args=generic_args, cb=self.__on_menu_add_remove_add_remove)
        self._index = -1

    # endregion Dunder Methods

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.PopupMenu",)

    # endregion Overrides

    # region Lazy Listeners

    def __on_menu_add_remove_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addMenuListener(self.events_listener_menu)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region MenuPartial Overrides
    def get_popup_menu(self, menu_id: int) -> Self | None:
        """
        Gets the popup menu from the menu item.
        """
        menu = self.component.getPopupMenu(menu_id)
        if menu is None:
            return None
        return self.__class__(menu)  # type: ignore

    # endregion MenuPartial Overrides

    # region MenuEvents Overrides
    if TYPE_CHECKING or mock_g.DOCS_BUILDING:

        def add_event_item_activated(self, cb: Callable[[Any, EventArgs, Self], None]) -> None:
            """
            Adds a callback for the item activated event for all menu items in the current instance.

            Args:
                cb (Callable[[Any, EventArgs, PopupMenuComp], None]): Callback function.
            """
            super().add_event_item_activated(cb)  # type: ignore

        def add_event_item_deactivated(self, cb: Callable[[Any, EventArgs, Self], None]) -> None:
            """
            Adds a callback for the item deactivated event for all menu items in the current instance.

            Args:
                cb (Callable[[Any, EventArgs, PopupMenuComp], None]): Callback function.
            """
            super().add_event_item_deactivated(cb)  # type: ignore

        def add_event_item_highlighted(self, cb: Callable[[Any, EventArgs, Self], None]) -> None:
            """
            Adds a callback for the item highlighted event for all menu items in the current instance.

            Args:
                cb (Callable[[Any, EventArgs, PopupMenuComp], None]): Callback function.
            """
            super().add_event_item_highlighted(cb)  # type: ignore

        def add_event_item_selected(self, cb: Callable[[Any, EventArgs, Self], None]) -> None:
            """
            Adds a callback for the item selected event for all menu items in the current instance.

            Args:
                cb (Callable[[Any, EventArgs, PopupMenuComp], None]): Callback function.
            """
            super().add_event_item_selected(cb)  # type: ignore

        def add_event_menu_events_disposing(self, cb: Callable[[Any, EventArgs, Self], None]) -> None:
            """
            Adds a callback for the menu events disposing event for all menu items in the current instance.

            Args:
                cb (Callable[[Any, EventArgs, PopupMenuComp], None]): Callback function.
            """
            super().add_event_menu_events_disposing(cb)  # type: ignore

        def remove_event_item_activated(self, cb: Callable[[Any, EventArgs, Self], None]) -> None:
            """
            Removes a callback for the item activated event for all menu items in the current instance.

            Args:
                cb (Callable[[Any, EventArgs, PopupMenuComp], None]): Callback function.
            """
            super().remove_event_item_activated(cb)  # type: ignore

        def remove_event_item_deactivated(self, cb: Callable[[Any, EventArgs, Self], None]) -> None:
            """
            Removes a callback for the item deactivated event for all menu items in the current instance.

            Args:
                cb (Callable[[Any, EventArgs, PopupMenuComp], None]): Callback function.
            """
            super().remove_event_item_deactivated(cb)  # type: ignore

        def remove_event_item_highlighted(self, cb: Callable[[Any, EventArgs, Self], None]) -> None:
            """
            Removes a callback for the item highlighted event for all menu items in the current instance.

            Args:
                cb (Callable[[Any, EventArgs, PopupMenuComp], None]): Callback function.
            """
            super().remove_event_item_highlighted(cb)  # type: ignore

        def remove_event_item_selected(self, cb: Callable[[Any, EventArgs, Self], None]) -> None:
            """
            Removes a callback for the item selected event for all menu items in the current instance.

            Args:
                cb (Callable[[Any, EventArgs, PopupMenuComp], None]): Callback function.
            """
            super().remove_event_item_selected(cb)  # type: ignore

        def remove_event_menu_events_disposing(self, cb: Callable[[Any, EventArgs, Self], None]) -> None:
            """
            Removes a callback for the menu events disposing event for all menu items in the current instance.

            Args:
                cb (Callable[[Any, EventArgs, PopupMenuComp], None]): Callback function.
            """
            super().remove_event_menu_events_disposing(cb)  # type: ignore

    # endregion MenuEvents Overrides

    # region Dunder Methods
    def __getitem__(self, key: int) -> dict:
        """Gets the item at the specified index."""
        result = {}
        menu_id = self.get_item_id(key)
        menu_type = self.get_item_type(key)
        result["id"] = menu_id
        result["type"] = menu_type
        if menu_type == MenuItemType.SEPARATOR:
            return result

        result["cmd"] = self.get_command(menu_id)
        result["help_cmd"] = self.get_command(menu_id)
        result["text"] = self.get_item_text(menu_id)
        result["tip_text"] = self.get_tip_help_text(menu_id)
        result["help"] = self.get_help_text(menu_id)
        result["is_enabled"] = self.is_item_enabled(menu_id)
        result["is_popup"] = self.is_popup_menu()
        return result

    def __iter__(self) -> Self:
        self._index = 0
        return self

    def __next__(self) -> int:
        """Gets the next item id."""
        if self._index >= self.get_item_count():
            self._index = -1
            raise StopIteration
        item_id = self.get_item_id(self._index)
        self._index += 1
        return item_id

    # endregion Dunder Methods

    # region Static Methods

    @classmethod
    def from_lo(cls, lo_inst: LoInst | None = None) -> Self:
        """
        Creates the instance from the Lo.

        Args:
            lo_inst (LoInst, optional): LoInst, Defaults to ``Lo.current_lo``.

        Returns:
            PopupMenuComp: The instance.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.loader import lo as mLo

        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        inst = lo_inst.create_instance_mcf(XPopupMenu, "com.sun.star.awt.PopupMenu", raise_err=True)  # type: ignore
        return cls(inst)

    # endregion Static Methods

    # region Add/Remove Events
    def subscribe_all_item_activated(self, cb: Callable[[Any, EventArgs, Self], None]) -> None:
        """
        Adds a callbacks for the item activated event for all menu and submenus items.

        Args:
            cb (Callable[[Any, EventArgs, PopupMenuComp], None]): Callback function.
        """

        def add_menu_event(mnu: PopupMenuComp) -> None:
            nonlocal cb
            with contextlib.suppress(Exception):
                mnu.add_event_item_activated(cb)  # type: ignore
            for i in mnu:
                pop_menu = mnu.get_popup_menu(i)
                if pop_menu is not None:
                    add_menu_event(pop_menu)

        add_menu_event(self)

    def subscribe_all_item_deactivated(self, cb: Callable[[Any, EventArgs, Self], None]) -> None:
        """
        Adds a callbacks for the item deactivated event for all menu and submenus items.

        Args:
            cb (Callable[[Any, EventArgs, PopupMenuComp], None]): Callback function.
        """

        def add_menu_event(mnu: PopupMenuComp) -> None:
            nonlocal cb
            with contextlib.suppress(Exception):
                mnu.add_event_item_deactivated(cb)  # type: ignore
            for i in mnu:
                pop_menu = mnu.get_popup_menu(i)
                if pop_menu is not None:
                    add_menu_event(pop_menu)

        add_menu_event(self)

    def subscribe_all_item_highlighted(self, cb: Callable[[Any, EventArgs, Self], None]) -> None:
        """
        Adds a callbacks for the item highlighted event for all menu and submenus items.

        Args:
            cb (Callable[[Any, EventArgs, PopupMenuComp], None]): Callback function.
        """

        def add_menu_event(mnu: PopupMenuComp) -> None:
            nonlocal cb
            with contextlib.suppress(Exception):
                mnu.add_event_item_highlighted(cb)  # type: ignore
            for i in mnu:
                pop_menu = mnu.get_popup_menu(i)
                if pop_menu is not None:
                    add_menu_event(pop_menu)

        add_menu_event(self)

    def subscribe_all_item_selected(self, cb: Callable[[Any, EventArgs, Self], None]) -> None:
        """
        Adds a callbacks for the item selected event for all menu and submenus items.

        Args:
            cb (Callable[[Any, EventArgs, PopupMenuComp], None]): Callback function.
        """

        def add_menu_event(mnu: PopupMenuComp) -> None:
            nonlocal cb
            with contextlib.suppress(Exception):
                mnu.add_event_item_selected(cb)  # type: ignore
            for i in mnu:
                pop_menu = mnu.get_popup_menu(i)
                if pop_menu is not None:
                    add_menu_event(pop_menu)

        add_menu_event(self)

    def unsubscribe_all_item_activated(self, cb: Callable[[Any, EventArgs, Self], None]) -> None:
        """
        Remove callbacks for the item selected event for all menu and submenus items.

        Args:
            cb (Callable[[Any, EventArgs, PopupMenuComp], None]): Callback function.
        """

        def remove_menu_event(mnu: PopupMenuComp) -> None:
            nonlocal cb
            with contextlib.suppress(Exception):
                mnu.remove_event_item_activated(cb)  # type: ignore
            for i in mnu:
                pop_menu = mnu.get_popup_menu(i)
                if pop_menu is not None:
                    remove_menu_event(pop_menu)

        remove_menu_event(self)

    def unsubscribe_all_item_deactivated(self, cb: Callable[[Any, EventArgs, Self], None]) -> None:
        """
        Remove callbacks for the item deactivated event for all menu and submenus items.

        Args:
            cb (Callable[[Any, EventArgs, PopupMenuComp], None]): Callback function.
        """

        def remove_menu_event(mnu: PopupMenuComp) -> None:
            nonlocal cb
            with contextlib.suppress(Exception):
                mnu.remove_event_item_deactivated(cb)  # type: ignore
            for i in mnu:
                pop_menu = mnu.get_popup_menu(i)
                if pop_menu is not None:
                    remove_menu_event(pop_menu)

        remove_menu_event(self)

    def unsubscribe_all_item_highlighted(self, cb: Callable[[Any, EventArgs, Self], None]) -> None:
        """
        Remove callbacks for the item highlighted event for all menu and submenus items.

        Args:
            cb (Callable[[Any, EventArgs, PopupMenuComp], None]): Callback function.
        """

        def remove_menu_event(mnu: PopupMenuComp) -> None:
            nonlocal cb
            with contextlib.suppress(Exception):
                mnu.remove_event_item_highlighted(cb)  # type: ignore
            for i in mnu:
                pop_menu = mnu.get_popup_menu(i)
                if pop_menu is not None:
                    remove_menu_event(pop_menu)

        remove_menu_event(self)

    def unsubscribe_all_item_selected(self, cb: Callable[[Any, EventArgs, Self], None]) -> None:
        """
        Remove callbacks for the item selected event for all menu and submenus items.

        Args:
            cb (Callable[[Any, EventArgs, PopupMenuComp], None]): Callback function.
        """

        def remove_menu_event(mnu: PopupMenuComp) -> None:
            nonlocal cb
            with contextlib.suppress(Exception):
                mnu.remove_event_item_selected(cb)  # type: ignore
            for i in mnu:
                pop_menu = mnu.get_popup_menu(i)
                if pop_menu is not None:
                    remove_menu_event(pop_menu)

        remove_menu_event(self)

    # endregion Add/Remove Events

    # region Properties

    @property
    def component(self) -> PopupMenu:
        """PopupMenu Component"""
        # overrides base class property
        # pylint: disable=no-member
        return cast("PopupMenu", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
