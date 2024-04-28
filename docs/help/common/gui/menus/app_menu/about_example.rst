.. _help_app_menu_about_example:

About Example for App Menu
==========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 1

This example demonstrates how to create a simple menu entry that displays an about dialog when clicked.

The :py:class:`~ooodev.gui.menu.menu_app.MenuApp` does execute commands directly.
This is different from the :ref:`help_popup_menu_about_example` example, which uses a popup menu and requires a callback to execute commands.

In the two examples below, the first example uses the :py:class:`~ooodev.gui.menu.ma.MACreator` class to create the menu entry.
and the second example uses the :py:class:`~ooodev.gui.menu.menu_app.MenuApp` class.

The :py:class:`~ooodev.gui.menu.ma.MACreator` class can also import and export Json data and can process more dynamic menu data.

Using MACreator
---------------

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        from typing import TYPE_CHECKING
        import uno

        from ooodev.calc import CalcDoc
        from ooodev.loader import Lo
        from ooodev.utils.kind.menu_lookup_kind import MenuLookupKind
        from ooodev.gui.menu.ma.ma_creator import MACreator

        if TYPE_CHECKING:
            pass

        def get_menu_data() -> list:
            new_menu = [{"Label": "About", "CommandURL": ".uno:About"}]
            return new_menu

        def main():
            loader = Lo.load_office(connector=Lo.ConnectPipe())
            doc = CalcDoc.create_doc(loader=loader, visible=True)
            try:
                sheet = doc.sheets[0]
                sheet[0, 0].value = "Hello, World!"

                tools_menu = doc.menu[MenuLookupKind.TOOLS]
                creator = MACreator()
                menu_data = creator.create(get_menu_data())

                for menu in menu_data[::-1]:
                    # insert the menus before the AutoComplete menu.
                    # loop in reverse to keep the order.
                    # save is set to false so that the menu is not saved to the configuration.
                    tools_menu.insert(menu=menu, after=".uno:AutoComplete", save=False)

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


Using Only MenuApp
------------------

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        from typing import TYPE_CHECKING
        import uno

        from ooodev.calc import CalcDoc
        from ooodev.loader import Lo
        from ooodev.utils.kind.menu_lookup_kind import MenuLookupKind

        if TYPE_CHECKING:
            pass

        def get_menu_data() -> list:
            new_menu = [{"Label": "About", "CommandURL": ".uno:About"}]
            return new_menu

        def main():
            loader = Lo.load_office(connector=Lo.ConnectPipe())
            doc = CalcDoc.create_doc(loader=loader, visible=True)
            try:
                sheet = doc.sheets[0]
                sheet[0, 0].value = "Hello, World!"

                tools_menu = doc.menu[MenuLookupKind.TOOLS]

                menu_data = get_menu_data()
                for menu in menu_data[::-1]:
                    # insert the menus before the AutoComplete menu.
                    # loop in reverse to keep the order.
                    # save is set to false so that the menu is not saved to the configuration.
                    tools_menu.insert(menu=menu, after=".uno:AutoComplete", save=False)

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

    .. _ca9505a8-71c4-489a-bfe2-60d77dea103e:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/ca9505a8-71c4-489a-bfe2-60d77dea103e
        :alt: Tools menu displaying about command.
        :figclass: align-center

        Tools menu displaying about command.