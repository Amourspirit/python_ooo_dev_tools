.. _help_menu_context_incept:

Intercepting a Context Menu
===========================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 3

Introduction
------------

The context menu is a popup menu that appears when you right-click on an object.
This example demonstrates how to intercept the context menu and add or modify menu items.

Incepting a Context Menu is a somewhat involved process. |odev| provides a much simpler way to intercept the context menu and add or modify menu items.

By subscribing to the notify context menu execute of the document view via  ``add_event_notify_context_menu_execute()`` method, you can intercept the context menu and add or modify menu items.
No need to implement the XContextMenuInterception_ interface and register the interceptor. This is all done for you.

The ``add_event_notify_context_menu_execute()`` method takes a callback function that is called when the context menu is about to be displayed.

The callback event is a ``EventArgsGeneric`` object that contains the event data. The event data is a :py:class:`~ooodev.adapter.ui.context_menu_interceptor_event_data.ContextMenuInterceptorEventData` object that contains the event and the action trigger container.

By checking the ``CommandURL`` of the first item in the action trigger container, you can determine the context menu that is about to be displayed. You can then add or modify menu items as needed.

In the example below we are looking to modify the context menu when a user right-clicks on a sheet tab of a Calc document.
We are looking for the ``.uno:Insert`` command URL at index ``0``. When found, a new menu item to the context menu.


Example Add Single Item
-----------------------

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        from typing import Any, TYPE_CHECKING
        import uno
        from ooo.dyn.ui.context_menu_interceptor_action import ContextMenuInterceptorAction as ContextMenuAction

        from ooodev.adapter.ui.context_menu_interceptor import ContextMenuInterceptor
        from ooodev.adapter.ui.context_menu_interceptor_event_data import ContextMenuInterceptorEventData
        from ooodev.calc import CalcDoc
        from ooodev.events.args.event_args_generic import EventArgsGeneric
        from ooodev.gui.menu.popup.action.action_trigger_container import ActionTriggerContainer
        from ooodev.gui.menu.popup.action.action_trigger_item import ActionTriggerItem
        from ooodev.loader import Lo

        
        if TYPE_CHECKING:
            from ooodev.calc.calc_sheet_view import CalcSheetView


        def on_menu_intercept(
            src: ContextMenuInterceptor,
            event: EventArgsGeneric[ContextMenuInterceptorEventData],
            view: CalcSheetView,
        ) -> None:
            try:
                # selection = event.event_data.event.selection.get_selection()
                # print(selection)
                container = event.event_data.event.action_trigger_container
                print(container[0].CommandURL)
                if container[0].CommandURL == ".uno:Insert":
                    container.insert_by_index(0, ActionTriggerItem(".uno:SelectTables", "Sheet..."))
                    event.event_data.action = ContextMenuAction.EXECUTE_MODIFIED


        def main():
            loader = Lo.load_office(connector=Lo.ConnectPipe())
            doc = CalcDoc.create_doc(loader=loader, visible=True)
            try:
                sheet = doc.sheets[0]
                sheet.set_active()
                sheet[0, 0].value = "Hello, World!"

                view = doc.get_view()
                view.add_event_notify_context_menu_execute(on_menu_intercept)
                # set a breakpoint here to see results.
                assert view

            finally:
                doc.close()
                Lo.close_office()


        if __name__ == "__main__":
            main()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Inserts a ``Sheet...`` menu command at the beginning of the popup. The result is see in :numref:`81cc6077-22c1-4077-a33d-be2b17cb391e`.

.. cssclass:: screen_shot

    .. _81cc6077-22c1-4077-a33d-be2b17cb391e:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/81cc6077-22c1-4077-a33d-be2b17cb391e
        :alt: Intercepting a Context Menu
        :figclass: align-center

        Add Sheet... to the top of Context Menu


Example Add Multiple Items
--------------------------

Alternatively Add a container (submenu).


.. tabs::

    .. code-tab:: python

        def on_menu_intercept(
            src: ContextMenuInterceptor,
            event: EventArgsGeneric[ContextMenuInterceptorEventData],
            view: CalcSheetView,
        ) -> None:
            try:
                container = event.event_data.event.action_trigger_container
                if container[0].CommandURL == ".uno:Insert":
                    items = ActionTriggerContainer()
                    items.insertByIndex(0, ActionTriggerItem(".uno:SelectTables", "Sheet..."))

                    item = ActionTriggerItem("GoTo", "Go to", sub_menu=items)
                    container.insert_by_index(7, item)
                    event.event_data.action = ContextMenuAction.EXECUTE_MODIFIED

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Inserts a ``Go To`` menu command with a submenu at index position of ``7`` of the popup. The result is see in :numref:`72c58518-81f0-452f-847d-5b68da29098f`.

.. cssclass:: screen_shot

    .. _72c58518-81f0-452f-847d-5b68da29098f:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/72c58518-81f0-452f-847d-5b68da29098f
        :alt: Intercepting a Context Menu
        :figclass: align-center

        Add Go to in the Context Menu

Related Topics
--------------

- :ref:`help_menu_context_incept_class_ex`
- :ref:`help_sample_context_menu_json_normal_data`
- :ref:`help_sample_context_menu_json_dynamic_data`

.. _XContextMenuInterception: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1ui_1_1XContextMenuInterception.html