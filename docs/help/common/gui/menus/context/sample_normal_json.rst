.. _help_sample_context_menu_json_normal_data:

Sample Context Menu JSON Normal Data
====================================

Sample JSON data for a Context menu. This data was outputted as normal format.

Mimics the popup menu in Calc when right-clicking on a cell.

.. seealso::

    - :ref:`help_menu_context_incept`

.. code-block:: json

    {
        "id": "ooodev.context_action_menu",
        "version": "0.41.0",
        "dynamic": false,
        "menus": [
            {
                "text": "~Cut",
                "command": ".uno:Cut"
            },
            {
                "text": "Cop~y",
                "command": ".uno:Copy"
            },
            {
                "text": "~Paste",
                "command": ".uno:Paste"
            },
            {
                "text": "Paste Special",
                "command": ".uno:PasteSpecialMenu",
                "submenu": [
                    {
                        "text": "Paste Unformatted Text",
                        "command": ".uno:PasteUnformatted"
                    },
                    {
                        "text": "-",
                        "separator_type": 0
                    },
                    {
                        "text": "My Paste Only Text",
                        "command": ".uno:PasteOnlyText"
                    },
                    {
                        "text": "Paste Only Text",
                        "command": ".uno:PasteOnlyValue"
                    },
                    {
                        "text": "Paste Only Formula",
                        "command": ".uno:PasteOnlyFormula"
                    },
                    {
                        "text": "-",
                        "separator_type": 0
                    },
                    {
                        "text": "Paste Transposed",
                        "command": ".uno:PasteTransposed"
                    },
                    {
                        "text": "-",
                        "separator_type": 0
                    },
                    {
                        "text": "Paste ~Special...",
                        "command": ".uno:PasteSpecial"
                    }
                ]
            },
            {
                "text": "-",
                "separator_type": 0
            },
            {
                "text": "Data Select",
                "command": ".uno:DataSelect"
            },
            {
                "text": "Current Validation",
                "command": ".uno:CurrentValidation"
            },
            {
                "text": "Define Current Name",
                "command": ".uno:DefineCurrentName"
            },
            {
                "text": "-",
                "separator_type": 0
            },
            {
                "text": "Insert cells",
                "command": ".uno:InsertCell"
            },
            {
                "text": "Del cells",
                "command": ".uno:DeleteCell"
            },
            {
                "text": "Delete",
                "command": ".uno:Delete"
            },
            {
                "text": "Merge Cells",
                "command": ".uno:MergeCells"
            },
            {
                "text": "Split Cell",
                "command": ".uno:SplitCell"
            },
            {
                "text": "-",
                "separator_type": 0
            },
            {
                "text": "Format Paintbrush",
                "command": ".uno:FormatPaintbrush"
            },
            {
                "text": "Reset Attributes",
                "command": ".uno:ResetAttributes"
            },
            {
                "text": "Format Styles Menu",
                "command": ".uno:FormatStylesMenu",
                "submenu": [
                    {
                        "text": "Edit Style",
                        "command": ".uno:EditStyle"
                    },
                    {
                        "text": "-",
                        "separator_type": 0
                    },
                    {
                        "text": "Default Cell Styles",
                        "command": ".uno:DefaultCellStylesmenu"
                    },
                    {
                        "text": "Accent1 Cell Styles",
                        "command": ".uno:Accent1CellStyles"
                    },
                    {
                        "text": "Accent2 Cell Styles",
                        "command": ".uno:"
                    },
                    {
                        "text": "Accent 3 Cell Styles",
                        "command": ".uno:Accent3CellStyles"
                    },
                    {
                        "text": "-",
                        "separator_type": 0
                    },
                    {
                        "text": "Bad Cell Styles",
                        "command": ".uno:BadCellStyles"
                    },
                    {
                        "text": "Error Cell Styles",
                        "command": ".uno:ErrorCellStyles"
                    },
                    {
                        "text": "Good Cell Styles",
                        "command": ".uno:GoodCellStyles"
                    },
                    {
                        "text": "Neutral Cell Styles",
                        "command": ".uno:NeutralCellStyles"
                    },
                    {
                        "text": "Warning Cell Styles",
                        "command": ".uno:WarningCellStyles"
                    },
                    {
                        "text": "-",
                        "separator_type": 0
                    },
                    {
                        "text": "Footnote Cell Styles",
                        "command": ".uno:FootnoteCellStyles"
                    },
                    {
                        "text": "Note Cell Styles",
                        "command": ".uno:NoteCellStyles"
                    }
                ]
            },
            {
                "text": "-",
                "separator_type": 0
            },
            {
                "text": "Insert Annotation",
                "command": ".uno:InsertAnnotation"
            },
            {
                "text": "Edit Annotation",
                "command": ".uno:EditAnnotation"
            },
            {
                "text": "Delete Note",
                "command": ".uno:DeleteNote"
            },
            {
                "text": "Show Note",
                "command": ".uno:ShowNote"
            },
            {
                "text": "Hide Note",
                "command": ".uno:HideNote"
            },
            {
                "text": "-",
                "separator_type": 0
            },
            {
                "text": "Format Sparkline",
                "command": ".uno:FormatSparklineMenu"
            },
            {
                "text": "-",
                "separator_type": 0
            },
            {
                "text": "Conditional Formatting...",
                "command": ".uno:CurrentConditionalFormatDialog"
            },
            {
                "text": "Current Conditional Format Manager Dialog ...",
                "command": ".uno:CurrentConditionalFormatManagerDialog"
            },
            {
                "text": "Format Cell Dialog ...",
                "command": ".uno:FormatCellDialog"
            }
        ]
    }