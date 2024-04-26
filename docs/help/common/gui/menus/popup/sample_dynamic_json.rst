.. _help_sample_popup_menu_json_dynamic_data:

Sample Popup Menu JSON Dynamic Data
===================================

Sample JSON data for a popup menu. This data was outputted as dynamic.

Mimics the popup menu in Calc when right-clicking on a cell.

.. code-block:: json

    {
        "id": "ooodev.popup_menu",
        "version": "0.41.0",
        "dynamic": true,
        "menus": [
            {
                "command": ".uno:Cut",
                "module": 20
            },
            {
                "command": ".uno:Copy",
                "module": 20
            },
            {
                "command": ".uno:Paste",
                "module": 20
            },
            {
                "text": "Paste Special",
                "command": ".uno:PasteSpecialMenu",
                "submenu": [
                    {
                        "command": ".uno:PasteUnformatted",
                        "module": 20
                    },
                    {
                        "text": "-"
                    },
                    {
                        "text": "My Paste Only Text",
                        "command": ".uno:PasteOnlyText",
                        "module": 0
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
                        "text": "-"
                    },
                    {
                        "text": "Paste Transposed",
                        "command": ".uno:PasteTransposed"
                    },
                    {
                        "text": "-"
                    },
                    {
                        "command": ".uno:PasteSpecial",
                        "module": 20
                    }
                ]
            },
            {
                "text": "-"
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
                "text": "-"
            },
            {
                "text": "Insert ~Cells...",
                "command": ".uno:InsertCell"
            },
            {
                "text": "Delete C~ells...",
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
                "text": "-"
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
                        "text": "-"
                    },
                    {
                        "text": "Default Cell Styles",
                        "command": ".uno:DefaultCellStylesmenu",
                        "style": 2
                    },
                    {
                        "text": "Accent1 Cell Styles",
                        "command": ".uno:Accent1CellStyles",
                        "style": 2
                    },
                    {
                        "text": "Accent2 Cell Styles",
                        "style": 2
                    },
                    {
                        "text": "Accent 3 Cell Styles",
                        "command": ".uno:Accent3CellStyles",
                        "style": 2
                    },
                    {
                        "text": "-"
                    },
                    {
                        "text": "Bad Cell Styles",
                        "command": ".uno:BadCellStyles",
                        "style": 2
                    },
                    {
                        "text": "Error Cell Styles",
                        "command": ".uno:ErrorCellStyles",
                        "style": 2
                    },
                    {
                        "text": "Good Cell Styles",
                        "command": ".uno:GoodCellStyles",
                        "style": 2
                    },
                    {
                        "text": "Neutral Cell Styles",
                        "command": ".uno:NeutralCellStyles",
                        "style": 2
                    },
                    {
                        "text": "Warning Cell Styles",
                        "command": ".uno:WarningCellStyles",
                        "style": 2
                    },
                    {
                        "text": "-"
                    },
                    {
                        "text": "Footnote Cell Styles",
                        "command": ".uno:FootnoteCellStyles",
                        "style": 2
                    },
                    {
                        "text": "Note Cell Styles",
                        "command": ".uno:NoteCellStyles",
                        "style": 2
                    }
                ]
            },
            {
                "text": "-"
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
                "text": "-"
            },
            {
                "text": "Format Sparkline",
                "command": ".uno:FormatSparklineMenu"
            },
            {
                "text": "-"
            },
            {
                "command": ".uno:CurrentConditionalFormatDialog",
                "module": 20
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