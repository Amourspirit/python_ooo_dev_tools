.. _help_menu_context_incept_class_ex:

Intercepting a Context Menu using a Class
=========================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 3

Introduction
------------

In this example, we will create a class that intercepts a context menu in a Calc document.
The class will add a new menu item to the sheet tab context menu.
The new menu item will be inserted after the "Insert" menu item.
This is not a practical example, but it demonstrates how to intercept a context menu.
Also demonstrates creating a new menu and saving and loading to and from a json file.

Logging is used to show the flow of the program.

Also note that is is possible to build a menu that can be used as data for the :py:meth:`ooodev.gui.menu.context.context_creator.ContextCreator.create` method using the :py:class:`~ooodev.gui.menu.popup.builder.BuilderItem` class.
See :ref:`help_popup_via_builder_item`.

Code Example
------------

The result of running the code example can be see in :numref:`4c8f660d-daa3-4e67-97e0-59f226085390`.

.. collapse:: Code Example
    :open:

    .. code-block:: python

        from __future__ import annotations
        from typing import TYPE_CHECKING
        import uno
        from ooo.dyn.awt.menu_item_style import MenuItemStyleEnum
        from ooo.dyn.ui.context_menu_interceptor_action import ContextMenuInterceptorAction
        from ooodev.adapter.ui.context_menu_interceptor import ContextMenuInterceptor
        from ooodev.adapter.ui.context_menu_interceptor_event_data import ContextMenuInterceptorEventData
        from ooodev.calc import CalcDoc
        from ooodev.events.args.event_args_generic import EventArgsGeneric
        from ooodev.gui.menu.context.action_trigger_item import ActionTriggerItem
        from ooodev.gui.menu.context.context_creator import ContextCreator
        from ooodev.loader import Lo
        from ooodev.utils.kind.module_names_kind import ModuleNamesKind
        from ooodev.io.log.named_logger import NamedLogger
        from ooodev.loader.inst.options import Options
        import logging

        if TYPE_CHECKING:
            from ooodev.calc.calc_sheet_view import CalcSheetView


        class CalcMenuIntercept:
            def __init__(self, doc: CalcDoc) -> None:
                self._fn_on_menu_intercept = self.on_menu_intercept
                self._match_cmd = ".uno:Insert"
                self._action_menu = None
                self._file_path = Lo.tmp_dir / "sheet_tab_menu_data.json"
                self._logger = NamedLogger(self.__class__.__name__)
                view = doc.get_view()
                view.add_event_notify_context_menu_execute(self._fn_on_menu_intercept)  # type: ignore

            def on_menu_intercept(
                self,
                src: ContextMenuInterceptor,
                event: EventArgsGeneric[ContextMenuInterceptorEventData],
                view: CalcSheetView,
            ) -> None:
                try:
                    # selection = event.event_data.event.selection.get_selection()
                    # print(selection)
                    container = event.event_data.event.action_trigger_container
                    if container[0].CommandURL == self._match_cmd:
                        self._logger.debug(f"Matched command: {self._match_cmd}")
                        items = self._get_menu()

                        item = ActionTriggerItem("GoTo", "Go to", sub_menu=items)
                        container.insert_by_index(7, item)  # type: ignore
                        event.event_data.action = ContextMenuInterceptorAction.EXECUTE_MODIFIED

                except Exception as e:
                    print(e)

            def _get_menu(self):
                if self._action_menu is None:
                    creator = ContextCreator()
                    data = self._load_menu_data()
                    self._action_menu = creator.create(data)
                else:
                    self._logger.debug("Returning existing menu")
                return self._action_menu

            def _load_menu_data(self):
                if self._file_path.exists():
                    self._logger.debug("Loading menu data from json file")
                    return self._get_menu_data_from_json()
                creator = ContextCreator()
                data = self._get_menu_data()
                self._logger.debug("Saving menu data to json file")
                creator.json_dump(file=self._file_path, menus=data)
                return data

            def _get_menu_data_from_json(self) -> list:
                return ContextCreator().json_load(self._file_path)

            def _get_menu_data(self) -> list:
                self._logger.debug("Creating new menu data")
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
                            {
                                "text": "Bad Cell Styles",
                                "command": ".uno:BadCellStyles",
                                "style": MenuItemStyleEnum.RADIOCHECK,
                            },
                            {
                                "text": "Error Cell Styles",
                                "command": ".uno:ErrorCellStyles",
                                "style": MenuItemStyleEnum.RADIOCHECK,
                            },
                            {
                                "text": "Good Cell Styles",
                                "command": ".uno:GoodCellStyles",
                                "style": MenuItemStyleEnum.RADIOCHECK,
                            },
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
                            {
                                "text": "Note Cell Styles",
                                "command": ".uno:NoteCellStyles",
                                "style": MenuItemStyleEnum.RADIOCHECK,
                            },
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
            loader = Lo.load_office(connector=Lo.ConnectPipe(), opt=Options(log_level=logging.DEBUG))
            doc = CalcDoc.create_doc(loader=loader, visible=True)
            try:
                menu_intercept = CalcMenuIntercept(doc)
                sheet = doc.sheets[0]
                sheet[0, 0].value = "Hello, World!"
                # set breakpoint here to see the menu
                assert menu_intercept

            finally:
                doc.close()
                Lo.close_office()


        if __name__ == "__main__":
            main()


New submenu added to context menu.

.. cssclass:: screen_shot

    .. _4c8f660d-daa3-4e67-97e0-59f226085390:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4c8f660d-daa3-4e67-97e0-59f226085390
        :alt: Context menu with new submenu
        :figclass: align-center
        :width: 550

        Context menu with new submenu.

Logging results
---------------

On first run the menu is created and saved to a json file.

.. code-block:: bash

    26/04/2024 16:16:26 - DEBUG - CalcMenuIntercept: Matched command: .uno:Insert
    26/04/2024 16:16:26 - DEBUG - CalcMenuIntercept: Creating new menu data
    26/04/2024 16:16:26 - DEBUG - CalcMenuIntercept: Saving menu data to json file
    26/04/2024 16:16:30 - DEBUG - CalcMenuIntercept: Matched command: .uno:Insert
    26/04/2024 16:16:30 - DEBUG - CalcMenuIntercept: Returning existing menu

On subsequent runs the menu is loaded from the json file.

.. code-block:: bash

    26/04/2024 16:17:37 - DEBUG - CalcMenuIntercept: Matched command: .uno:Insert
    26/04/2024 16:17:37 - DEBUG - CalcMenuIntercept: Loading menu data from json file
    26/04/2024 16:17:40 - DEBUG - CalcMenuIntercept: Matched command: .uno:Insert
    26/04/2024 16:17:40 - DEBUG - CalcMenuIntercept: Returning existing menu


Related Topics
--------------

- :ref:`help_menu_context_incept`
- :ref:`help_sample_context_menu_json_normal_data`
- :ref:`help_sample_context_menu_json_dynamic_data`