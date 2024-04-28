.. _help_menubar_menu_about_example:

About Example for MenuBar
=========================


.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 1

This example demonstrates how to add to an existing popup menu that displays an about dialog when clicked.

Popup menus do not process the command directly. This is unlike the :ref:`help_app_menu_about_example`.
Instead, Popup menus can raise an event that can be used to process the command.
In this example, the event is processed by the ``on_tools_menu_select`` function.
This function is called when a menu item on the tools popup menu is selected.
The function checks if the command matches ``.uno:About`` and then function executes the command.

Example Code
------------

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        from typing import Any, cast, TYPE_CHECKING
        import uno

        from ooodev.calc import CalcDoc
        from ooodev.loader import Lo
        from ooodev.utils.kind.menu_lookup_kind import MenuLookupKind
        from ooodev.gui.menu.menu_bar import MenuBar

        if TYPE_CHECKING:
            from com.sun.star.awt import MenuEvent
            from ooodev.events.args.event_args import EventArgs
            from ooodev.gui.menu.popup_menu import PopupMenu

        def on_tools_menu_select(src: Any, event: EventArgs, menu: PopupMenu) -> None:
            me = cast("MenuEvent", event.event_data)
            command = menu.get_command(me.MenuId)
            if command == ".uno:About":
                menu.execute_cmd(command)

        def get_menu_bar(doc: CalcDoc) -> MenuBar:
            # get the menubar of the active document
            doc.activate()
            comp = doc.get_frame_comp()
            if comp is None:
                raise ValueError("No frame component found")
            lm = comp.layout_manager
            mb = lm.get_menu_bar()
            if mb is None:
                raise ValueError("No menu bar found")
            return mb

        def main():
            loader = Lo.load_office(connector=Lo.ConnectPipe())
            doc = CalcDoc.create_doc(loader=loader, visible=True)
            try:
                sheet = doc.sheets[0]
                sheet[0, 0].value = "Hello, World!"

                mb = get_menu_bar(doc)
                menu_id, _ = mb.find_item_menu_id(str(MenuLookupKind.TOOLS))
                menu = mb.get_popup_menu(menu_id)
                if menu is None:
                    raise ValueError("No Tools Menu found")

                next_id = menu.insert_item_after(menu_id=-1, text="About", after=".uno:AutoComplete")
                menu.set_command(menu_id=next_id, command=".uno:About")
                menu.add_event_item_selected(on_tools_menu_select)

                # set breakpoint here to see the menu.
                assert True
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

When you run the example, an About menu entry is displayed on the tools menu.
When you click the About command, the about dialog is displayed.

.. cssclass:: screen_shot

    .. _ca9505a8-71c4-489a-bfe2-60d77dea103e_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/ca9505a8-71c4-489a-bfe2-60d77dea103e
        :alt: Tools menu displaying about command.
        :figclass: align-center

        Tools menu displaying about command.