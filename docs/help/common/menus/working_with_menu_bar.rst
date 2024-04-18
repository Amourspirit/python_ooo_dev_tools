.. _help_working_with_menu_bar:

Working with MenuBar
====================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 1

Working with :ref:`ooodev.gui.menu.MenuBar`.

Introduction
------------

The :ref:`ooodev.gui.menu.MenuBar` differs from menu app is several ways.
The :ref:`ooodev.gui.menu.MenuApp` is best suited for adding and removing menus although it is not limited to those functions.

The :ref:`ooodev.gui.menu.MenuApp` and :ref:`ooodev.gui.menu.PopupMenu` offer functionality for all Popup menus on the menu bar, also allows for subscribing to menu events.

Getting access to MenuBar
-------------------------

The menu bar for the document window is access via the Layout Manager.


.. tabs::

    .. code-tab:: python

        from ooodev.calc import CalcDoc
        # ...

        doc = CalcDoc.from_current_doc()

        # to get the frame the document must be active
        # activation would not be necessary if the document is already active or running in a macro.
        doc.activate()
        frame_component = doc.get_frame_comp()
        assert frame_component is not None
        mb = frame_component.layout_manager.get_menu_bar()
        assert mb is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Accessing Menu
--------------

Accessing are root menu is done in the following way:

.. tabs::

    .. code-tab:: python

        from ooodev.utils.kind.menu_lookup_kind import MenuLookupKind
        # ...

        menu_id, _ = mb.find_item_menu_id(cmd=str(MenuLookupKind.TOOLS))
        tools_popup = mb.get_popup_menu(menu_id)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The :py:class:`~ooodev.utils.kind.menu_lookup_kind.MenuLookupKind` is for convenience and in this case returns ``.uno:ToolsMenu``.
In this example the ``tools_popup`` is an instance of :ref:`ooodev.gui.menu.PopupMenu`.

The :py:meth:`MenuBar.find_item_menu_id() <ooodev.gui.menu.MenuBar.MenuBar.find_item_menu_id>` and :py:meth:`MenuBar.find_item_pos() <ooodev.gui.menu.MenuBar.MenuBar.find_item_pos>` methods return a tuple of two elements.
The first element is the menu id for ``find_item_menu_id()`` and the zero based index position for ``find_item_pos()``,
the second element is the popup menu that the search command was found on.
The element is only useful only applies the ``search_sub_menu=True`` is set. More on this later.

Accessing a popup menu of an instance of :ref:`ooodev.gui.menu.PopupMenu` is done in the following way:

.. tabs::

    .. code-tab:: python

        >>> menu_id, _ = tools_popup.find_item_menu_id(".uno:ToolsFormsMenu")
        >>> forms_menu = tools_popup.get_popup_menu(menu_id)
        >>> len(forms_menu)
        12

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Searching for a Popup Menu
--------------------------

Recursively search
^^^^^^^^^^^^^^^^^^

Recursively search MenuBar
""""""""""""""""""""""""""

A :ref:`ooodev.gui.menu.MenuBar` instance can be recursively searched by setting ``search_sub_menu=True``
In this can we are finding ``Tools -> Forms -> Design Mode`` command.

.. tabs::

    .. code-tab:: python

        >>> design_menu_id, popup = mb.find_item_menu_id(
        >>> 	cmd=".uno:SwitchControlDesignMode",
        >>> 	search_sub_menu=True,
        >>> )
        >>> if popup is not None:
        >>> 	print(popup.get_command(design_menu_id), design_menu_id, sep=": ")
        .uno:SwitchControlDesignMode: 563

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Recursively search PopupMenu
""""""""""""""""""""""""""""

A :ref:`ooodev.gui.menu.PopupMenu` instance can be recursively searched by setting ``search_sub_menu=True``

.. tabs::

    .. code-tab:: python

        >>> tool_menu_id, _ = mb.find_item_menu_id(str(MenuLookupKind.TOOLS))
        >>> tool_popup = mb.get_popup_menu(tool_menu_id)
        >>> assert tool_popup is not None
        >>> design_menu_id, popup = tool_popup.find_item_menu_id(
        >>> 	cmd=".uno:SwitchControlDesignMode",
        >>> 	search_sub_menu=True,
        >>> )
        >>> if popup is not None:
        >>> 	print(popup.get_command(design_menu_id), design_menu_id, sep=": ")
        .uno:SwitchControlDesignMode: 563

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Caching
-------

Both :ref:`ooodev.gui.menu.MenuBar` and :ref:`ooodev.gui.menu.PopupMenu` have a built in caching.
The caching is provided by the :py:class:`ooodev.utils.lru_cache.LRUCache` class.
Both have a ``cache`` property that allow for cache to be modified.
The caching speed up searching by caching found result.
The is means with the same object is search more the once then after the first search the result is pulled from the cache.
If Needed the cache for a object can be cleared as follows:

.. tabs::

    .. code-tab:: python

        doc.activate()
        comp = doc.get_frame_comp()
        assert comp is not None
        lm = comp.layout_manager
        mb = lm.get_menu_bar()
        assert mb is not None

        # ...
        # clear the cache
        mb.cache.clear()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Turn cache off can be done by setting ``capacity=0``.

.. tabs::

    .. code-tab:: python

        i, command_menu = mb.find_item_menu_id("MyCommand", True)
        if command_menu:
            # turn cache off ofr menu
            command_menu.cache.capacity = 0

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Listening to Events
-------------------

Listening to events for any or all menu popup menus is possible. The following Event can be listened for:

- item activated via :py:meth:`~ooodev.gui.menu.PopupMenu.add_event_item_activated` or :py:meth:`~ooodev.gui.menu.PopupMenu.subscribe_all_item_activated`
- item highlighted  via :py:meth:`~ooodev.gui.menu.PopupMenu.add_event_item_highlighted` or :py:meth:`~ooodev.gui.menu.PopupMenu.subscribe_all_item_highlighted`
- item selected via :py:meth:`~ooodev.gui.menu.PopupMenu.add_event_item_selected` or :py:meth:`~ooodev.gui.menu.PopupMenu.subscribe_all_item_selected`
- item deactivated via :py:meth:`~ooodev.gui.menu.PopupMenu.add_event_item_deactivated` or :py:meth:`~ooodev.gui.menu.PopupMenu.subscribe_all_item_deactivated`

There are also corresponding ``remove_*`` and ``unsubscribe_*`` events.
All ``add_event_*`` and ``remove_*``  methods  work on the current ``PopupMenu``.
All ``subscribe_*``  and ``unsubscribe_*`` methods work on the current ``PopupMenu`` and any child ``PopupMenu`` recursively.

General Example
^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        from ooodev.gui.menu.popup_menu import PopupMenu
        from typing import Any, cast, TYPE_CHECKING

        if TYPE_CHECKING:
            from com.sun.star.awt import MenuEvent

        def on_my_cmd_menu_select(src: Any, event: EventArgs, menu: PopupMenu) -> None:
            print("Menu Selected")
            me = cast("MenuEvent", event.event_data)
            print("MenuId", me.MenuId)

            pos = menu.get_item_pos(me.MenuId)
            menu_type = menu.get_item_type(pos)
            if menu_type == MenuItemType.SEPARATOR:
                return
            enabled = menu.is_item_enabled(me.MenuId)
            if not enabled:
                return
            cmd = menu.get_command(me.MenuId)
            if cmd == "MyCommand1":
                print("Found Command", cmd)
                doc = Lo.current_doc
                doc.msgbox("Found Command 1", title="Info", boxtype=1)
            elif cmd == "MyCommand2":
                print("Found Command", cmd)

        def on_menu_my_cmd_highlighted(src: Any, event: EventArgs, menu: PopupMenu) -> None:
            print("Menu Highlighted")
            me = cast("MenuEvent", event.event_data)
            print("MenuId", me.MenuId)
            pos = menu.get_item_pos(me.MenuId)
            menu_type = menu.get_item_type(pos)
            if menu_type == MenuItemType.SEPARATOR:
                return
            cmd = menu.get_command(me.MenuId)
            if cmd:
                print("Command", cmd)

        def create_cmd_menu() -> PopupMenu:
            pm = PopupMenu.from_lo()
            pm.insert_item(0, "~First Entry", MenuItemStyleKind.NONE, 0)
            pm.insert_item(1, "~First Radio Entry", MenuItemStyleKind.RADIOCHECK | MenuItemStyleKind.AUTOCHECK, 1)
            pm.insert_item(2, "~Second Radio Entry", MenuItemStyleKind.RADIOCHECK | MenuItemStyleKind.AUTOCHECK, 2)
            pm.insert_item(3, "~Third RadioEntry", MenuItemStyleKind.RADIOCHECK | MenuItemStyleKind.AUTOCHECK, 3)
            pm.insert_separator(4)
            pm.insert_item(4, "F~ifth Entry", MenuItemStyleKind.CHECKABLE | MenuItemStyleKind.AUTOCHECK, 5)
            pm.insert_item(5, "~Fourth Entry", MenuItemStyleKind.CHECKABLE | MenuItemStyleKind.AUTOCHECK, 6)
            pm.enable_item(1, False)
            pm.insert_item(6, "~Sixth Entry", 0, 7)
            pm.insert_item(7, "~EightEntry", MenuItemStyleKind.RADIOCHECK | MenuItemStyleKind.AUTOCHECK, 8)
            for i in range(8):
                pm.set_command(i, f"MyCommand{i}")
            pm.check_item(2, True)
            return pm

        def add_cmd_menu() -> None:
            menu_id, _ = mb.find_item_menu_id(str(MenuLookupKind.TOOLS))
            if menu_id < 0:
                raise Exception("Tools Menu not found!")

            tools_popup = mb.get_popup_menu(menu_id)
            if not tools_popup:
                raise Exception(f"Did not find popup menu for menu id: {menu_id}")

            new_id = tools_popup.get_max_menu_id() + 1
            pm = create_cmd_menu()
            # add a new Tools Entry as the first menu item in the Tools Menu
            te_popup.insert_item(new_id, "~Tools Entry", MenuItemStyleKind.NONE, 0)
            te_popup.set_command(new_id, "MyCommand")

            
            te_popup.set_popup_menu(new_id, pm)
            new_pop = te_popup.get_popup_menu(new_id)
            if new_pop is None:
                # should never happen but also keeps type chekers happy
                rasie Exception(f"Not able to find menu inserted into Tools Entry menu for menu idL {new_id})
            
            # subscribe all menu items to highlighted and selected
            mb.subscribe_all_item_highlighted(on_menu_my_cmd_highlighted)
            mb.subscribe_all_item_selected(on_my_cmd_menu_select)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

In the above example code a ``create_cmd_menu()`` method is called by ``add_cmd_menu()`` method.

The ``create_cmd_menu()`` creates a new popup menu, ``add_cmd_menu()`` finds the ``Tools`` Popup Menu and inserts a new ``Tools Entry`` menu item as the first entry.
Next the popup menu is assigned to the ``Tools Entry`` and finally the ``Tools Entry`` menu item and its popup menu subscribe to events.

Class Example
^^^^^^^^^^^^^

Example of capturing menu id's and taking action based on menu id in a menu event.

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        from typing import Any, cast, TYPE_CHECKING
        import logging
        from ooodev.calc import CalcDoc
        from ooodev.events.args.event_args import EventArgs
        from ooodev.gui.menu.menu_bar import MenuBar
        from ooodev.gui.menu.popup_menu import PopupMenu
        from ooodev.loader import Lo
        from ooodev.loader.inst.options import Options
        from ooodev.macro.script.macro_script import MacroScript
        from ooodev.utils.kind.menu_item_style_kind import MenuItemStyleKind
        from ooodev.utils.kind.menu_lookup_kind import MenuLookupKind

        if TYPE_CHECKING:
            from com.sun.star.awt import MenuEvent

        MY_MENU = None
        # keep class instance alive.


        class MyMenu:
            """Sample Menu Class"""
            def __init__(self, doc: CalcDoc) -> None:
                self._doc = doc
                self._exec_menu_ids = set()
                self._init_events()
                self._add_tools_entry()

            def _init_events(self) -> None:
                # capture class method so it can be used like a function for callbacks.
                self._fn_on_my_cmd_menu_select = self._on_my_cmd_menu_select

            def _on_my_cmd_menu_select(self, src: Any, event: EventArgs, menu: PopupMenu) -> None:
                # manage menu events
                me = cast("MenuEvent", event.event_data)
                if me.MenuId in self._exec_menu_ids:
                    menu.execute_cmd(me.MenuId, in_thread=True)
                    return
                cmd = menu.get_command(me.MenuId)
                if cmd:
                    self._doc.msgbox(f"Command: {cmd}", title="Info", boxtype=1)

            def _get_menu_bar(self) -> MenuBar:
                # get the menubar of the active document
                self._doc.activate()
                comp = self._doc.get_frame_comp()
                if comp is None:
                    raise ValueError("No frame component found")
                lm = comp.layout_manager
                mb = lm.get_menu_bar()
                if mb is None:
                    raise ValueError("No menu bar found")
                return mb

            def _get_tools_popup(self) -> PopupMenu:
                # get the Tools Popup Menu from the tool bar.
                mb = self._get_menu_bar()
                menu_id, _ = mb.find_item_menu_id(str(MenuLookupKind.TOOLS))
                tools_popup = mb.get_popup_menu(menu_id)
                if tools_popup is None:
                    raise ValueError("No tools popup found")
                return tools_popup

            def _add_tools_entry(self) -> None:
                # add a new entry to the Tools menu with a popup menu.
                mnu_command = "MyCommand"
                tools_popup = self._get_tools_popup()
                menu_id, _ = tools_popup.find_item_menu_id(mnu_command)
                if menu_id != -1:
                    raise ValueError("Menu already exists")
                new_id = tools_popup.get_max_menu_id() + 1
                tools_popup.insert_item(new_id, "~Tools Entry", MenuItemStyleKind.NONE, 0)
                tools_popup.set_command(new_id, mnu_command)

                pm = self._get_tools_popup_menu()
                tools_popup.set_popup_menu(new_id, pm)

                new_pop = tools_popup.get_popup_menu(new_id)
                if new_pop is None:
                    raise ValueError("No new popup found")
                new_pop.add_event_item_selected(self._fn_on_my_cmd_menu_select)

            def _get_tools_popup_menu(self) -> PopupMenu:
                # create a new popup menu for the Tools menu entry.
                pm = PopupMenu.from_lo()
                pm.insert_item(0, "~Toggle Formula", MenuItemStyleKind.NONE, 0)
                pm.insert_separator(1)
                pm.insert_item(2, "~Hello World", MenuItemStyleKind.NONE, 2)
                pm.insert_separator(3)
                pm.insert_item(4, "~Other Entry 1", MenuItemStyleKind.NONE, 4)
                pm.insert_item(5, "~Other Entry 2", MenuItemStyleKind.NONE, 5)

                pm.set_command(0, ".uno:ToggleFormula")
                self._exec_menu_ids.add(0)

                url = MacroScript.get_url_script(
                    name="HelloWorldPython", library="HelloWorld", language="Python", location="share"
                )
                pm.set_command(2, url)
                self._exec_menu_ids.add(2)

                pm.set_command(4, "MyCommand1")
                pm.set_command(5, "MyCommand2")

                return pm


        def main():
            global MY_MENU
            loader = Lo.load_office(connector=Lo.ConnectPipe(), opt=Options(log_level=logging.DEBUG, lo_cache_size=400))
            doc = CalcDoc.create_doc(loader=loader, visible=True)
            try:

                if MY_MENU is None:
                    MY_MENU = MyMenu(doc)
                assert MY_MENU is not None

                Lo.delay(20_000)
            # ...
            finally:
                doc.close()
                Lo.close_office()


        if __name__ == "__main__":
            main()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The above class example shows one way of capturing Menu id's and then taking action when the menu item is clicked.

.. note::

    When menus are first added using :ref:`ooodev.gui.menu.MenuApp` then the menu commands would be automatically executed by LibreOffice.

When the ``Toggle Formula`` menu it clicked will toggle formula display for the sheet via the ``.uno:ToggleFormula`` dispatch command.
The :py:class:`~ooodev.macro.script.MacroScript` class is used to create a  URL to an existing python script,
``vnd.sun.star.script:HelloWorld.py$HelloWorldPython?language=Python&location=share``.
The URL is assigned to the ``Hello World`` menu.
When the ``Hello World`` menu item is clicked it runs a python script that open a new Writer document and writes ``Hello World (in Python)`` into it.

In this example the menu id's that are to be executed are stored in `self._exec_menu_ids`. Otherwise the command displays a message box.

.. tabs::

    .. code-tab:: python

        pm.set_command(0, ".uno:ToggleFormula")
        self._exec_menu_ids.add(0)
        # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

And ``_exec_menu_ids`` is checked in the event callback.

.. tabs::

    .. code-tab:: python

        def _on_my_cmd_menu_select(self, src: Any, event: EventArgs, menu: PopupMenu) -> None:
            # manage menu events
            me = cast("MenuEvent", event.event_data)
            if me.MenuId in self._exec_menu_ids:
                menu.execute_cmd(me.MenuId, in_thread=True)
                return
            # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Intercept Menus
^^^^^^^^^^^^^^^

For Context menus there is a `XContextMenuInterceptor <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1ui_1_1XContextMenuInterceptor.html>`__ that can be used to incept menus and change them.
There does not seem to be an equivalent for Menu Bar Popup Menus.

This means means when a menu item is a build in dispatch command will not be possible to intercept the menu and cancel the event like a content menu.
A much lower level dispatch listener would need to be implemented for this.
This seems to be complex approach and I have not seen any clear solution for intercepting or listening to all dispatch commands.

The work around for this seems to be either set a custom command that is not a ``.uno:some_cmd`` dispatch command and
then listen to your custom event as done in the example above or set the menu command to a macro and handle the rest in the macro.

Other Notes
-----------

I found debugging searching for menus to be unpredictable. At Least in Windows 10.
When debugging recursive search on Menus such as the following:


.. tabs::

    .. code-tab:: python

        doc.activate()
        comp = doc.get_frame_comp()
        assert comp is not None
        lm = comp.layout_manager
        mb = lm.get_menu_bar()
        assert mb is not None

        i, m = mb.find_item_menu_id("MyCommand", True)
        j, val = m.find_item_menu_id("MyCommand1", True)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

I would get unpredictable results, often the menu item was not found.
I the GUI of Calc many menus would be come unavailable (greyed out).
I ran the code without putting VS Code into debug mode then the code works fine.
Not sure what may have caused this.

Related Topics
--------------

- :ref:`help_creating_menu_using_menu_app`
- :ref:`help_working_with_menu_app`