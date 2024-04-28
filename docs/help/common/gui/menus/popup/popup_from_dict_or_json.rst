.. _help_popup_from_dict_or_json:

Popup from Dictionary or JSON
=============================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 3


Demonstrates creating a popup menu using a dictionary.
This example mimics the popup menu displayed Calc Cell.

This example is the same data that is use to export :ref:`help_sample_popup_menu_json_normal_data` and :ref:`help_sample_popup_menu_json_dynamic_data`.

The ``module`` key is used to instruct :py:class:`~ooodev.gui.menu.popup.PopupCreator` to lookup up the command info from :py:class:`~ooodev.gui.commands.CmdInfo`
The :py:class:`~ooodev.utils.kind.module_names_kind.ModuleNamesKind` is used to specify the module to lookup the command info from.

Example
-------

.. collapse:: Example code
    :open:

    .. tabs::

        .. code-tab:: python

            from __future__ import annotations
            from typing import Any, cast, TYPE_CHECKING
            import uno
            from com.sun.star.awt import Rectangle
            from ooo.dyn.awt.menu_item_style import MenuItemStyleEnum

            from ooodev.calc import CalcDoc
            from ooodev.loader import Lo
            from ooodev.events.args.event_args import EventArgs
            from ooodev.gui.menu.popup.popup_creator import PopupCreator
            from ooodev.utils.kind.module_names_kind import ModuleNamesKind

            if TYPE_CHECKING:
                from com.sun.star.awt import MenuEvent
                from ooodev.gui.menu.popup_menu import PopupMenu


            def on_menu_select(src: Any, event: EventArgs, menu: PopupMenu) -> None:
                print("Menu Selected")
                me = cast("MenuEvent", event.event_data)
                print("MenuId", me.MenuId)
                command = menu.get_command(me.MenuId)
                if command:
                    print("Command", command)
                    # check if command is a dispatch command
                    if menu.is_dispatch_cmd(command):
                        menu.execute_cmd(command)

            def get_popup_menu() -> list:

                new_menu = [
                    {"command": ".uno:Cut", "module": ModuleNamesKind.SPREADSHEET_DOCUMENT},
                    {"command": ".uno:Copy", "module": ModuleNamesKind.SPREADSHEET_DOCUMENT},
                    {"command": ".uno:Paste", "module": ModuleNamesKind.SPREADSHEET_DOCUMENT},
                    {
                        "text": "Paste Special",
                        "command": ".uno:PasteSpecialMenu",
                        "submenu": [
                            {
                                # "text": "Paste Unformatted",
                                "command": ".uno:PasteUnformatted",
                                "module": ModuleNamesKind.SPREADSHEET_DOCUMENT,
                            },
                            {"text": "-"},
                            {"text": "My Paste Only Text", "command": ".uno:PasteOnlyText", "module": ModuleNamesKind.NONE},
                            {"text": "Paste Only Text", "command": ".uno:PasteOnlyValue"},
                            {"text": "Paste Only Formula", "command": ".uno:PasteOnlyFormula"},
                            {"text": "-"},
                            {"text": "Paste Transposed", "command": ".uno:PasteTransposed"},
                            {"text": "-"},
                            {
                                "command": ".uno:PasteSpecial",
                                "module": ModuleNamesKind.SPREADSHEET_DOCUMENT,
                            },
                        ],
                    },
                    {"text": "-"},
                    {"text": "Data Select", "command": ".uno:DataSelect"},
                    {"text": "Current Validation", "command": ".uno:CurrentValidation"},
                    {"text": "Define Current Name", "command": ".uno:DefineCurrentName"},
                    {"text": "-"},
                    {"text": "Insert cells", "command": ".uno:InsertCell"},
                    {"text": "Del cells", "command": ".uno:DeleteCell"},
                    {"text": "Delete", "command": ".uno:Delete"},
                    {"text": "Merge Cells", "command": ".uno:MergeCells"},
                    {"text": "Split Cell", "command": ".uno:SplitCell"},
                    {"text": "-"},
                    {"text": "Format Paintbrush", "command": ".uno:FormatPaintbrush"},
                    {"text": "Reset Attributes", "command": ".uno:ResetAttributes"},
                    {
                        "text": "Format Styles Menu",
                        "command": ".uno:FormatStylesMenu",
                        "submenu": [
                            {"text": "Edit Style", "command": ".uno:EditStyle"},
                            {"text": "-"},
                            {
                                "text": "Default Cell Styles",
                                "command": ".uno:DefaultCellStylesmenu",
                                "style": MenuItemStyleEnum.RADIOCHECK,
                            },
                            {
                                "text": "Accent1 Cell Styles",
                                "command": ".uno:Accent1CellStyles",
                                "style": MenuItemStyleEnum.RADIOCHECK,
                            },
                            {
                                "text": "Accent2 Cell Styles",
                                "style": MenuItemStyleEnum.RADIOCHECK,
                            },
                            {
                                "text": "Accent 3 Cell Styles",
                                "command": ".uno:Accent3CellStyles",
                                "style": MenuItemStyleEnum.RADIOCHECK,
                            },
                            {"text": "-"},
                            {"text": "Bad Cell Styles", "command": ".uno:BadCellStyles", "style": MenuItemStyleEnum.RADIOCHECK},
                            {
                                "text": "Error Cell Styles",
                                "command": ".uno:ErrorCellStyles",
                                "style": MenuItemStyleEnum.RADIOCHECK,
                            },
                            {"text": "Good Cell Styles", "command": ".uno:GoodCellStyles", "style": MenuItemStyleEnum.RADIOCHECK},
                            {
                                "text": "Neutral Cell Styles",
                                "command": ".uno:NeutralCellStyles",
                                "style": MenuItemStyleEnum.RADIOCHECK,
                            },
                            {
                                "text": "Warning Cell Styles",
                                "command": ".uno:WarningCellStyles",
                                "style": MenuItemStyleEnum.RADIOCHECK,
                            },
                            {
                                "text": "-",
                            },
                            {
                                "text": "Footnote Cell Styles",
                                "command": ".uno:FootnoteCellStyles",
                                "style": MenuItemStyleEnum.RADIOCHECK,
                            },
                            {"text": "Note Cell Styles", "command": ".uno:NoteCellStyles", "style": MenuItemStyleEnum.RADIOCHECK},
                        ],
                    },
                    {"text": "-"},
                    {"text": "Insert Annotation", "command": ".uno:InsertAnnotation"},
                    {"text": "Edit Annotation", "command": ".uno:EditAnnotation"},
                    {"text": "Delete Note", "command": ".uno:DeleteNote"},
                    {"text": "Show Note", "command": ".uno:ShowNote"},
                    {"text": "Hide Note", "command": ".uno:HideNote"},
                    {"text": "-"},
                    {"text": "Format Sparkline", "command": ".uno:FormatSparklineMenu"},
                    {"text": "-"},
                    {"command": ".uno:CurrentConditionalFormatDialog", "module": ModuleNamesKind.SPREADSHEET_DOCUMENT},
                    {
                        "text": "Current Conditional Format Manager Dialog ...",
                        "command": ".uno:CurrentConditionalFormatManagerDialog",
                    },
                    {"text": "Format Cell Dialog ...", "command": ".uno:FormatCellDialog"},
                ]
                return new_menu

            def main():
                loader = Lo.load_office(connector=Lo.ConnectPipe())
                doc = CalcDoc.create_doc(loader=loader, visible=True)
                try:
                    creator = PopupCreator()
                    menus = get_popup_menu()
                    pm = creator.create(menus)
                    pm.subscribe_all_item_selected(on_menu_select)
                    rect = Rectangle(100, 100, 100, 100)
                    doc.activate()
                    pm.execute(doc.get_frame().ComponentWindow, rect, 0)
                    # place a breakpoint here to inspect the menu
                    assert pm
                finally:
                    doc.close()
                    Lo.close_office()


            if __name__ == "__main__":
                main()

        .. only:: html

            .. cssclass:: tab-none

                .. group-tab:: None

Command Values
--------------

The :py:class:`~ooodev.gui.commands.CmdInfo` class (see :ref:`help_getting_info_on_commands`) makes it possible to look up command information.
There are few ways to do this.

Auto Command Values
^^^^^^^^^^^^^^^^^^^

On way is to just use the build in ability of the :py:class:`~ooodev.gui.menu.popup.PopupCreator` class.


.. tabs::

    .. code-tab:: python

        # auto command entry.
        {"command": ".uno:Cut", "module": ModuleNamesKind.SPREADSHEET_DOCUMENT},

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Auto command entries are created by including the ``module`` key in a menu entry.
This instructs :py:class:`~ooodev.gui.menu.popup.PopupCreator` to get the information by looking up the command information using :py:class:`~ooodev.gui.commands.CmdInfo` to fill in other popup information.
Note not every command has an entry in the sources that :py:class:`~ooodev.gui.commands.CmdInfo` pull from.

The Command Data for ``.uno:Cut`` is as follows:

.. tabs::

    .. code-tab:: python

        CmdData(
            command='.uno:Copy',
            label='Cop~y',
            name='Copy',
            popup=False,
            properties=1,
            popup_label='',
            tooltip_label='',
            target_url='',
            is_experimental=False,
            module_hotkey='',
            global_hotkey='Ctrl+C'
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Manual Command Lookup
^^^^^^^^^^^^^^^^^^^^^

It is also possible to lookup command info manually. A few modification to the example code:

.. tabs::

    .. code-tab:: python

        def get_cmd_data(cmd: str, mod_kind: str | ModuleNamesKind) -> CmdData | None:
            # CmdInfo() is a singleton.
            return CmdInfo().get_cmd_data(mode_name=mod_kind, cmd=cmd)


        def get_calc_command_text(cmd: str, default: str) -> str:
            cmd_data = get_cmd_data(cmd, ModuleNamesKind.SPREADSHEET_DOCUMENT)
            if cmd_data is not None:
                return cmd_data.label or cmd_data.name
            else:
                return default

        def main():
            # ...
            creator = PopupCreator()
            menus = get_popup_menu()
            pm = creator.create(menus)
            # other code

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Entries in the menu can now use the ``get_calc_command_text()`` method to lookup names for commands.

.. tabs::

    .. code-tab:: python

            {"text": get_calc_command_text(".uno:InsertCell", "Insert cells"), "command": ".uno:InsertCell"},

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. _help_popup_from_dict_or_json_events:

Events
------

Menu Creation Events
^^^^^^^^^^^^^^^^^^^^

Events are triggered while the menu is being built. Optionally these events can be subscribe to that can modify the menu creation.
Menu entries can have a ``data`` key.

.. tabs::

    .. code-tab:: python

            {"text": "Del cells", "command": ".uno:DeleteCell", "data": "hook_me_up"},

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The ``data`` key is strictly for use in the event callbacks, for the developer to use at their discretion.
The ``on_popup_created()`` can be used to subscribe to popup menu events. See the next section for more.

.. tabs::

    .. code-tab:: python

        def on_menu_select(src: Any, event: EventArgs, menu: PopupMenu) -> None:
            print("Menu Selected")
            me = cast("MenuEvent", event.event_data)
            print("MenuId", me.MenuId)
            command = menu.get_command(me.MenuId)
            if command:
                print("Command", command)
                # check if command is a dispatch command
                if menu.is_dispatch_cmd(command):
                    menu.execute_cmd(command)

        def on_popup_created(src, event: EventArgs):
            # print(f"on_before_process: {event.event_data}")
            e_data = cast(dict, event.event_data)
            popup_menu = cast("PopupMenu", e_data["popup_menu"])
            popup_menu.add_event_item_selected(on_menu_select)

        def on_after_process(src, event: EventArgs):
            e_data = cast(dict, event.event_data)
            popup_item = cast("PopupItem", e_data["popup_item"])
            popup_menu = cast("PopupMenu", e_data["popup_menu"])
            if popup_item.data == "hook_me_up":
                print(f"on_after_process: {popup_menu}")
                print("Hooked up!")


        def main():
            loader = Lo.load_office(connector=Lo.ConnectPipe())
            doc = CalcDoc.create_doc(loader=loader, visible=True)
            try:
                creator = PopupCreator()
                creator.subscribe_after_process(on_after_process)
                creator.subscribe_popup_created(on_popup_created)

                menus = get_popup_menu()
                pm = creator.create(menus)
                # ...
            finally:
                doc.close()
                Lo.close_office()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

No Text Callback

When a menu entry is using the ``module`` entry and the text for the entry is not found for the command then an event is raises.
It is possible to subscribe to these events and manually provide a value if needed.

The callback ``event_data`` is a dictionary with keys:

- ``module_kind``: :py:class:`~ooodev.utils.kind.module_names_kind.ModuleNamesKind`
- ``cmd``: Command as a string.
- ``index``: Index as an integer.
- ``menu``: Menu Data as a dictionary.


.. tabs::

    .. code-tab:: python

        def on_no_module_text(src, event: CancelEventArgs):
            # print(f"on_before_process: {event.event_data}")
            e_data = cast(dict, event.event_data)
            menu_data = cast(dict, e_data["menu"])
            if "data" in menu_data:
                # assign the data tag as the menu text.
                menu_data["text"] = str(menu_data["data"])
            else:
                # skip adding this menu item.
                event.cancel = True

        def main():
            loader = Lo.load_office(connector=Lo.ConnectPipe())
            doc = CalcDoc.create_doc(loader=loader, visible=True)
            try:
                creator = PopupCreator()
                creator.subscribe_module_no_text(on_no_module_text)

                menus = get_popup_menu()
                pm = creator.create(menus)
                # ...
            finally:
                doc.close()
                Lo.close_office()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Popup Menu Events
^^^^^^^^^^^^^^^^^

There are a few ways to subscribe to popup menu events.


During Creation
"""""""""""""""

In this example the ``on_popup_created()`` is called when a new popup menu has been created.
It is used to subscribe the popup to the ``on_menu_select()`` callback.

.. tabs::

    .. code-tab:: python

        def on_menu_select(src: Any, event: EventArgs, menu: PopupMenu) -> None:
            print("Menu Selected")
            me = cast("MenuEvent", event.event_data)
            print("MenuId", me.MenuId)
            command = menu.get_command(me.MenuId)
            if command:
                print("Command", command)
                # check if command is a dispatch command
                if menu.is_dispatch_cmd(command):
                    menu.execute_cmd(command)

        def on_popup_created(src, event: EventArgs):
            # print(f"on_before_process: {event.event_data}")
            e_data = cast(dict, event.event_data)
            popup_menu = cast("PopupMenu", e_data["popup_menu"])
            popup_menu.add_event_item_selected(on_menu_select)



        def main():
            loader = Lo.load_office(connector=Lo.ConnectPipe())
            doc = CalcDoc.create_doc(loader=loader, visible=True)
            try:
                creator = PopupCreator()
                creator.subscribe_popup_created(on_popup_created)

                menus = get_popup_menu()
                pm = creator.create(menus)
                # ...
            finally:
                doc.close()
                Lo.close_office()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

After Creation
""""""""""""""

It is also possible to subscribe to callbacks after the the popup creation is finished.

Single Popup
~~~~~~~~~~~~

If a popup has no sub menus:

.. tabs::

    .. code-tab:: python

        def main():
            # ...
            creator = PopupCreator()

            menus = get_popup_menu()
            pm = creator.create(menus)
            pm.add_event_item_activated(on_menu_select)
            # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Popup with sub menu
~~~~~~~~~~~~~~~~~~~

.. tabs::

    .. code-tab:: python

        def main():
            # ...
            creator = PopupCreator()

            menus = get_popup_menu()
            pm = creator.create(menus)
            pm.subscribe_all_item_selected(on_menu_select)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. _help_popup_from_dict_or_json_import_export:

Export/Import Json
------------------

A menu can be loaded and saved from json data.

The json data must have root attribute of ``id`` that has a value of ``ooodev.popup_menu`` to be considered valid.
The root attribute ``version`` is optional and is the ``version`` of OooDev that the menus was created with.

Saving
^^^^^^

Save as Dynamic
""""""""""""""""

Saves the JSON seen in :ref:`help_sample_popup_menu_json_dynamic_data`.

.. tabs::

    .. code-tab:: python

        def main():
            # ...
            creator = PopupCreator()
            menus = get_popup_menu()
            json_str = creator.json_dumps(menus, dynamic=True)
            with open("popup_menu.json", "w") as f:
                f.write(json_str)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Save as Normal
""""""""""""""

Save the JSON see in :ref:`help_sample_popup_menu_json_normal_data`.

.. tabs::

    .. code-tab:: python

        def main():
            # ...
            creator = PopupCreator()
            menus = get_popup_menu()
            json_str = creator.json_dumps(menus, dynamic=False)
            with open("popup_menu.json", "w") as f:
                f.write(json_str)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Load JSON
^^^^^^^^^

Load File
"""""""""


.. tabs::

    .. code-tab:: python

        def main():
            # ...
            creator = PopupCreator()
            menus = PopupCreator.json_load("popup_menu.json")
            pm = creator.create(menus)

            pm.subscribe_all_item_selected(on_menu_select)
            rect = Rectangle(100, 100, 100, 100)
            doc.activate()
            pm.execute(doc.get_frame().ComponentWindow, rect, 0)
            # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Load String
"""""""""""

.. tabs::

    .. code-tab:: python

        def main():
            # ...
            creator = PopupCreator()
            json_str = get_json_str()
            menus = PopupCreator.json_loads(json_str)
            pm = creator.create(menus)

            pm.subscribe_all_item_selected(on_menu_select)
            rect = Rectangle(100, 100, 100, 100)
            doc.activate()
            pm.execute(doc.get_frame().ComponentWindow, rect, 0)
            # ...

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None