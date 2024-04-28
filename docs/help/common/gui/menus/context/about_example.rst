.. _help_context_menu_about_example:

About Example for Context Menu
==============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 1

This example demonstrates how to create a simple menu entry that displays an about dialog when clicked.

Menu entries added to a context are not saved to the configuration. They are only available for the current session.
The menu entries are automatically executed when clicked. Adding a callback is not possible for context menu items.

This is different from the :ref:`help_popup_menu_about_example` example, which uses a popup menu and requires a callback to execute commands.

In the two examples below, the first example uses the :py:class:`~ooodev.gui.menu.context.context_creator.ContextCreator` class to create the menu entry.
and the second example uses the :py:class:`~ooodev.gui.menu.context.action_trigger_item.ActionTriggerItem` class.

The :py:class:`~ooodev.gui.menu.context.context_creator.ContextCreator` class can also import and export Json data and can process more dynamic menu data.

Note that ``event.event_data.action = ContextMenuAction.EXECUTE_MODIFIED`` is required to notify the context menu that it has been modified.

Using MACreator
---------------

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        from typing import Any
        import contextlib
        import uno
        from ooo.dyn.ui.context_menu_interceptor_action import ContextMenuInterceptorAction as ContextMenuAction

        from ooodev.adapter.ui.context_menu_interceptor import ContextMenuInterceptor
        from ooodev.adapter.ui.context_menu_interceptor_event_data import ContextMenuInterceptorEventData
        from ooodev.calc import CalcDoc
        from ooodev.events.args.event_args_generic import EventArgsGeneric
        from ooodev.gui.menu.context.context_creator import ContextCreator
        from ooodev.loader import Lo

        def on_menu_intercept(
            src: ContextMenuInterceptor, event: EventArgsGeneric[ContextMenuInterceptorEventData], view: Any
        ) -> None:
            try:
                container = event.event_data.event.action_trigger_container
                with contextlib.suppress(Exception):
                    # don't block other menus if there is an issue.
                    # check the first and last items in the container
                    if container[0].CommandURL == ".uno:Cut" and container[-1].CommandURL == ".uno:FormatCellDialog":
                        index = container.get_command_index(".uno:DeleteCell")

                        menu_container = get_context_menu()
                        container.insert_by_index(index + 1, menu_container[0])  # type: ignore

                        event.event_data.action = ContextMenuAction.EXECUTE_MODIFIED

            except Exception as e:
                print(e)

        def get_context_menu():
            # ContextCreator can load a menu from a JSON file or
            # create a menu with submenus from a list of dictionaries.
            creator = ContextCreator()
            return creator.create([{"text": "About", "command": ".uno:About"}])

        def main():
            loader = Lo.load_office(connector=Lo.ConnectPipe())
            doc = CalcDoc.create_doc(loader=loader, visible=True)
            try:
                sheet = doc.sheets[0]
                sheet.set_active()
                sheet[0, 0].value = "Hello, World!"

                view = doc.get_view()
                view.add_event_notify_context_menu_execute(on_menu_intercept)
                assert view

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
        from typing import Any
        import contextlib
        import uno
        from ooo.dyn.ui.context_menu_interceptor_action import ContextMenuInterceptorAction as ContextMenuAction

        from ooodev.adapter.ui.context_menu_interceptor import ContextMenuInterceptor
        from ooodev.adapter.ui.context_menu_interceptor_event_data import ContextMenuInterceptorEventData
        from ooodev.calc import CalcDoc
        from ooodev.events.args.event_args_generic import EventArgsGeneric
        from ooodev.gui.menu.context.action_trigger_item import ActionTriggerItem
        from ooodev.loader import Lo

        def on_menu_intercept(
            src: ContextMenuInterceptor, event: EventArgsGeneric[ContextMenuInterceptorEventData], view: Any
        ) -> None:
            try:
                container = event.event_data.event.action_trigger_container
                with contextlib.suppress(Exception):
                    # don't block other menus if there is an issue.
                    # check the first and last items in the container
                    if container[0].CommandURL == ".uno:Cut" and container[-1].CommandURL == ".uno:FormatCellDialog":
                        # for i, itm in enumerate(container):
                        #     if container.is_separator(itm):
                        #         continue
                        #     print(f"{i}: {itm.CommandURL}")
                        index = container.get_command_index(".uno:DeleteCell")

                        # sheet cell context menu
                        container.insert_by_index(index + 1, ActionTriggerItem(".uno:About", "About"))  # type: ignore

                        event.event_data.action = ContextMenuAction.EXECUTE_MODIFIED

            except Exception as e:
                print(e)

        def main():
            loader = Lo.load_office(connector=Lo.ConnectPipe())
            doc = CalcDoc.create_doc(loader=loader, visible=True)
            try:
                sheet = doc.sheets[0]
                sheet.set_active()
                sheet[0, 0].value = "Hello, World!"

                view = doc.get_view()
                view.add_event_notify_context_menu_execute(on_menu_intercept)
                assert view

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

    .. _4f7fd91e-4a6d-46ef-8f5d-f3a12a61094a:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4f7fd91e-4a6d-46ef-8f5d-f3a12a61094a
        :alt: Context menu displaying about command.
        :figclass: align-center

        Context menu displaying about command.