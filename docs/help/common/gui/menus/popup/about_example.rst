.. _help_popup_menu_about_example:

About Example for Popup
=======================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 1

This example demonstrates how to create a simple popup menu that displays an about dialog when clicked.

Popup menus do not process the command directly. This is unlike the :ref:`help_app_menu_about_example`.
Instead, Popup menus can raise an event that can be used to process the command.
In this example, the event is processed by the ``on_menu_select`` function.
This function is called when a menu item is selected.
The function checks if the command is a dispatch command.
If it is, the function executes the command.

In the two examples below, the first example uses the :py:class:`~ooodev.gui.menu.popup.PopupCreator` class to create the popup menu,
and the second example uses the :py:class:`~ooodev.gui.menu.PopupMenu` class.

If you just need a simple popup menu, you can use the :py:class:`~ooodev.gui.menu.PopupMenu` class.
:py:class:`~ooodev.gui.menu.popup.PopupCreator` is useful when you need to create a more complex popup menu and supports importing and exporting JSON files.


Using PopupCreator
------------------

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        from typing import Any, cast, TYPE_CHECKING
        import uno
        from com.sun.star.awt import Rectangle

        from ooodev.calc import CalcDoc
        from ooodev.gui.menu.popup.popup_creator import PopupCreator
        from ooodev.loader import Lo

        if TYPE_CHECKING:
            from com.sun.star.awt import MenuEvent
            from ooodev.events.args.event_args import EventArgs
            from ooodev.gui.menu.popup_menu import PopupMenu

        def on_menu_select(src: Any, event: EventArgs, menu: PopupMenu) -> None:
            print("Menu Selected")
            me = cast("MenuEvent", event.event_data)
            command = menu.get_command(me.MenuId)
            if command:
                # check if command is a dispatch command
                if menu.is_dispatch_cmd(command):
                    menu.execute_cmd(command)

        def get_popup_menu() -> list:
            new_menu = [{"text": "About", "command": ".uno:About"}]
            return new_menu

        def main():
            loader = Lo.load_office(connector=Lo.ConnectPipe())
            doc = CalcDoc.create_doc(loader=loader, visible=True)
            try:

                creator = PopupCreator()
                pm = creator.create(get_popup_menu())
                pm.add_event_item_selected(on_menu_select)
                rect = Rectangle(100, 100, 100, 100)
                doc.activate()
                pm.execute(doc.get_frame().ComponentWindow, rect, 0)
                # set a breakpoint here to see the popup menu.
                assert pm
            finally:
                doc.close()
                Lo.close_office()


        if __name__ == "__main__":
            main()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Using PopupMenu
---------------

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        from typing import Any, cast, TYPE_CHECKING
        import uno
        from com.sun.star.awt import Rectangle

        from ooodev.calc import CalcDoc
        from ooodev.gui.menu.popup_menu import PopupMenu
        from ooodev.loader import Lo

        if TYPE_CHECKING:
            from com.sun.star.awt import MenuEvent
            from ooodev.events.args.event_args import EventArgs

        def on_menu_select(src: Any, event: EventArgs, menu: PopupMenu) -> None:
            print("Menu Selected")
            me = cast("MenuEvent", event.event_data)
            command = menu.get_command(me.MenuId)
            if command:
                # check if command is a dispatch command
                if menu.is_dispatch_cmd(command):
                    menu.execute_cmd(command)

        def main():
            loader = Lo.load_office(connector=Lo.ConnectPipe())
            doc = CalcDoc.create_doc(loader=loader, visible=True)
            try:

                pm = PopupMenu.from_lo()
                pm.insert_item(0, "About", 0, 0)
                pm.set_command(0, ".uno:About")
                pm.add_event_item_selected(on_menu_select)

                rect = Rectangle(100, 100, 100, 100)
                doc.activate()
                pm.execute(doc.get_frame().ComponentWindow, rect, 0)
                # set a breakpoint here to see the popup menu.
                assert pm
            finally:
                doc.close()
                Lo.close_office()

        if __name__ == "__main__":
            main()


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Output
------

When you run the example, a popup menu is displayed.
When you click the About command, the about dialog is displayed.

.. cssclass:: screen_shot

    .. _8501ffdc-2b85-41fa-aca7-4a68b05494b2:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/8501ffdc-2b85-41fa-aca7-4a68b05494b2
        :alt: Popup menu displaying about command.
        :figclass: align-center

        Popup menu displaying about command.