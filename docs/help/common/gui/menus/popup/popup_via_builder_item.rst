.. _help_popup_via_builder_item:

Popup using BuilderItem
=======================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 3


The :py:class:`~ooodev.gui.menu.popup.builder.BuilderItem` class is a class that allows you to create a popup menu in an object oriented way.
It is possible to create a popup menu using a dictionary or JSON as well. See :ref:`help_popup_from_dict_or_json` for more information.

The result of the example method could also be used to export Json or to create a popup menu. See :ref:`help_popup_from_dict_or_json_import_export`.

.. tabs::

    .. code-tab:: python

        import uno
        from ooo.dyn.awt.menu_item_style import MenuItemStyleEnum
        from ooodev.utils.kind.module_names_kind import ModuleNamesKind
        from ooodev.gui.menu.popup.builder.builder_item import BuilderItem as BI
        from ooodev.gui.menu.popup.builder.sep_item import SepItem as SI

        def get_popup_menu() -> list:
            result = []
            sep = SI()
            pop = BI(command=".uno:Cut", module=ModuleNamesKind.SPREADSHEET_DOCUMENT)
            result.append(pop.to_dict())

            pop = BI(command=".uno:Copy", module=ModuleNamesKind.SPREADSHEET_DOCUMENT)
            result.append(pop.to_dict())

            pop = BI(text="Paste Special", command=".uno:PasteSpecialMenu")
            pop.submenu_add_popup(command=".uno:PasteUnformatted", module=ModuleNamesKind.SPREADSHEET_DOCUMENT)
            pop.submenu_add_separator()
            pop.submenu_add_popup(text="My Paste Only Text", command=".uno:PasteOnlyText", module=ModuleNamesKind.NONE)
            pop.submenu_add_popup(text="Paste Only Text", command=".uno:PasteOnlyValue")
            # pop.submenu_add_popup(text="Paste Only Formula", command=".uno:PasteOnlyFormula") Same as next line
            pop.submenu.append(BI(text="Paste Only Formula", command=".uno:PasteOnlyFormula"))
            pop.submenu_add_separator()
            pop.submenu_add_popup(command=".uno:PasteSpecial", module=ModuleNamesKind.SPREADSHEET_DOCUMENT)
            result.append(pop.to_dict())

            result.append(sep.to_dict())

            pop = BI(text="Data Select", command=".uno:DataSelect")
            result.append(pop.to_dict())
            pop = BI(text="Current Validation", command=".uno:CurrentValidation")
            result.append(pop.to_dict())
            pop = BI(text="Define Current Name", command=".uno:DefineCurrentName")
            result.append(pop.to_dict())
            result.append(sep.to_dict())
            pop = BI(text="Insert cells", command=".uno:InsertCell")
            result.append(pop.to_dict())
            pop = BI(text="Del cells", command=".uno:DeleteCell")
            result.append(pop.to_dict())
            pop = BI(text="Delete", command=".uno:Delete")
            result.append(pop.to_dict())
            pop = BI(text="Merge Cells", command=".uno:MergeCells")
            result.append(pop.to_dict())
            pop = BI(text="Split Cell", command=".uno:SplitCell")
            result.append(pop.to_dict())
            result.append(sep.to_dict())
            pop = BI(text="Format Paintbrush", command=".uno:FormatPaintbrush")
            result.append(pop.to_dict())
            pop = BI(text="Reset Attributes", command=".uno:ResetAttributes")
            result.append(pop.to_dict())

            pop = BI(text="Format Styles Menu", command=".uno:FormatStylesMenu")
            pop.submenu_add_popup(text="Edit Style", command=".uno:EditStyle")
            pop.submenu_add_separator()
            pop.submenu_add_popup(
                text="Default Cell Styles", command=".uno:DefaultCellStylesmenu", style=MenuItemStyleEnum.RADIOCHECK
            )
            pop.submenu_add_popup(
                text="Accent1 Cell Styles", command=".uno:Accent1CellStyles", style=MenuItemStyleEnum.RADIOCHECK
            )
            pop.submenu_add_popup(text="Accent2 Cell Styles", style=MenuItemStyleEnum.RADIOCHECK)
            pop.submenu_add_popup(
                text="Accent 3 Cell Styles", command=".uno:Accent3CellStyles", style=MenuItemStyleEnum.RADIOCHECK
            )
            pop.submenu_add_separator()
            pop.submenu_add_popup(text="Bad Cell Styles", command=".uno:BadCellStyles", style=MenuItemStyleEnum.RADIOCHECK)
            pop.submenu_add_popup(text="Error Cell Styles", command=".uno:ErrorCellStyles", style=MenuItemStyleEnum.RADIOCHECK)
            pop.submenu_add_popup(text="Good Cell Styles", command=".uno:GoodCellStyles", style=MenuItemStyleEnum.RADIOCHECK)
            pop.submenu_add_popup(
                text="Neutral Cell Styles", command=".uno:NeutralCellStyles", style=MenuItemStyleEnum.RADIOCHECK
            )
            pop.submenu_add_popup(
                text="Warning Cell Styles", command=".uno:WarningCellStyles", style=MenuItemStyleEnum.RADIOCHECK
            )
            pop.submenu_add_separator()
            pop.submenu_add_popup(
                text="Footnote Cell Styles", command=".uno:FootnoteCellStyles", style=MenuItemStyleEnum.RADIOCHECK
            )
            pop.submenu_add_popup(text="Note Cell Styles", command=".uno:NoteCellStyles", style=MenuItemStyleEnum.RADIOCHECK)
            result.append(pop.to_dict())

            result.append(sep.to_dict())
            pop = BI(text="Insert Annotation", command=".uno:InsertAnnotation")
            result.append(pop.to_dict())

            pop = BI(text="Edit Annotation", command=".uno:EditAnnotation")
            result.append(pop.to_dict())

            pop = BI(text="Delete Note", command=".uno:DeleteNote")
            result.append(pop.to_dict())

            pop = BI(text="Show Note", command=".uno:ShowNote")
            result.append(pop.to_dict())

            pop = BI(text="Hide Note", command=".uno:HideNote")
            result.append(pop.to_dict())

            result.append(sep.to_dict())

            pop = BI(text="Format Sparkline", command=".uno:FormatSparklineMenu")
            result.append(pop.to_dict())

            result.append(sep.to_dict())
            pop = BI(command=".uno:CurrentConditionalFormatDialog", module=ModuleNamesKind.SPREADSHEET_DOCUMENT)
            result.append(pop.to_dict())

            pop = BI(
                text="Current Conditional Format Manager Dialog ...", command=".uno:CurrentConditionalFormatManagerDialog"
            )
            result.append(pop.to_dict())

            pop = BI(text="Format Cell Dialog ...", command=".uno:FormatCellDialog")
            result.append(pop.to_dict())

            return result

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None
