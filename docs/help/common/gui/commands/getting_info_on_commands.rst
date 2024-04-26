.. _help_getting_info_on_commands:

Getting Info on Commands
=========================

LibreOffice has Thousands of Dispatch Commands, many can bee seen on The `Wiki <https://wiki.documentfoundation.org/Development/DispatchCommands>`__.

Getting information on a command is relatively easy using the :py:class:`~ooodev.gui.commands.CmdInfo` class.

The Commands are split into categories. The :py:class:`~ooodev.utils.kind.module_names_kind.ModuleNamesKind` enum can be used to access the different categories.

.. tabs::

    .. code-tab:: python

        from __future__ import annotations

        from ooodev.calc import CalcDoc
        from ooodev.loader import Lo
        from ooodev.gui.commands import CmdInfo
        from ooodev.utils.kind.module_names_kind import ModuleNamesKind


        def main():
            loader = Lo.load_office(connector=Lo.ConnectPipe())
            doc = CalcDoc.create_doc(loader=loader, visible=False)
            try:

                start_time = time.time()
                inst = CmdInfo() # singleton class

                cmd_data = inst.get_cmd_data(ModuleNamesKind.SPREADSHEET_DOCUMENT, ".uno:Copy")
                if cmd_data:
                    print(cmd_data)

            finally:
                doc.close()
                Lo.close_office()


        if __name__ == "__main__":
            main()

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Running the above module will output:

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


The ``global_hotkey`` or ``module_hotkey`` may delimited by ``;`` if there is more then one shortcut keys.

This snippet shows how you might get the short keys.

.. tabs::

    .. code-tab:: python

        from ooodev.utils.string.str_list import StrList
        from ooodev.gui.menu import Shortcuts

        # ...
        class PopupProcessor(EventsPartial):
            # ...
            def _process_shortcut(self, pop: PopupItem) -> None:
                """Process shortcut"""
                keys = pop.shortcut.strip()
                if not keys:
                    return
                sl_keys = StrList.from_str(keys)
                for key in sl_keys:
                    if not key:
                        continue
                    kv = Shortcuts.to_key_event(key)
                    if kv is not None:
                        self._popup.set_accelerator_key_event(pop.menu_id, kv)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Note that gathering up command data is a rather involved process.
The :py:class:`~ooodev.gui.commands.CmdInfo` class takes advantage of the  :py:class:`~ooodev.utils.cache.file_cache.PickleCache` and the :py:class:`~ooodev.utils.cache.LRUCache` to cache the various categories are they are found. This make other searches very fast by comparison.

It is also possible to search across all categories for a command using ``find_command()`` method.
Thanks to the caching a query like this takes about ``1/10th`` of a second on an average computer.

.. tabs::

    .. code-tab:: python

        import pprint
        # ...
        inst = CmdInfo()
        data = inst.find_command(".uno:Copy")
        pprint.pprint(data)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Outputs data for all modules.

.. tabs::

    .. code-tab:: python

        {'com.sun.star.chart2.ChartDocument': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='Ctrl+C', global_hotkey='Ctrl+C')],
        'com.sun.star.drawing.DrawingDocument': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='Ctrl+C', global_hotkey='Ctrl+C')],
        'com.sun.star.formula.FormulaProperties': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='', global_hotkey='Ctrl+C')],
        'com.sun.star.frame.Bibliography': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='', global_hotkey='Ctrl+C')],
        'com.sun.star.frame.StartModule': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='Ctrl+C', global_hotkey='Ctrl+C')],
        'com.sun.star.presentation.PresentationDocument': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='Ctrl+C', global_hotkey='Ctrl+C')],
        'com.sun.star.report.ReportDefinition': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='', global_hotkey='Ctrl+C')],
        'com.sun.star.script.BasicIDE': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='', global_hotkey='Ctrl+C')],
        'com.sun.star.sdb.DataSourceBrowser': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='', global_hotkey='Ctrl+C')],
        'com.sun.star.sdb.FormDesign': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='', global_hotkey='Ctrl+C')],
        'com.sun.star.sdb.OfficeDatabaseDocument': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='', global_hotkey='Ctrl+C')],
        'com.sun.star.sdb.QueryDesign': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='', global_hotkey='Ctrl+C')],
        'com.sun.star.sdb.RelationDesign': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='', global_hotkey='Ctrl+C')],
        'com.sun.star.sdb.TableDataView': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='', global_hotkey='Ctrl+C')],
        'com.sun.star.sdb.TableDesign': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='', global_hotkey='Ctrl+C')],
        'com.sun.star.sdb.TextReportDesign': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='', global_hotkey='Ctrl+C')],
        'com.sun.star.sdb.ViewDesign': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='', global_hotkey='Ctrl+C')],
        'com.sun.star.sheet.SpreadsheetDocument': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='', global_hotkey='Ctrl+C')],
        'com.sun.star.text.GlobalDocument': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='', global_hotkey='Ctrl+C')],
        'com.sun.star.text.TextDocument': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='', global_hotkey='Ctrl+C')],
        'com.sun.star.text.WebDocument': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='', global_hotkey='Ctrl+C')],
        'com.sun.star.xforms.XMLFormDocument': [CmdData(command='.uno:Copy', label='Cop~y', name='Copy', popup=False, properties=1, popup_label='', tooltip_label='', target_url='', is_experimental=False, module_hotkey='', global_hotkey='Ctrl+C')]}

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None