.. _help_sample_popup_menu_json_normal_data:

Sample Popup Menu JSON Normal Data
==================================

Sample JSON data for a popup menu. This data was outputted as normal format.

Mimics the popup menu in Calc when right-clicking on a cell.

.. code-block:: json

    {
        "id": "ooodev.popup_menu",
        "version": "0.41.0",
        "dynamic": false,
        "menus": [
            {
                "text": "~Cut",
                "command": ".uno:Cut",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": "Ctrl+X"
            },
            {
                "text": "Cop~y",
                "command": ".uno:Copy",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": "Ctrl+C"
            },
            {
                "text": "~Paste",
                "command": ".uno:Paste",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": "Ctrl+V"
            },
            {
                "text": "Paste Special",
                "command": ".uno:PasteSpecialMenu",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": "",
                "submenu": [
                    {
                        "text": "Paste Unformatted Text",
                        "command": ".uno:PasteUnformatted",
                        "style": 0,
                        "checked": false,
                        "enabled": true,
                        "default": false,
                        "help_command": "",
                        "help_text": "",
                        "tip_help_text": "",
                        "shortcut": "Shift+Ctrl+Alt+V"
                    },
                    {
                        "text": "-"
                    },
                    {
                        "text": "My Paste Only Text",
                        "command": ".uno:PasteOnlyText",
                        "style": 0,
                        "checked": false,
                        "enabled": true,
                        "default": false,
                        "help_command": "",
                        "help_text": "",
                        "tip_help_text": "",
                        "shortcut": ""
                    },
                    {
                        "text": "Paste Only Text",
                        "command": ".uno:PasteOnlyValue",
                        "style": 0,
                        "checked": false,
                        "enabled": true,
                        "default": false,
                        "help_command": "",
                        "help_text": "",
                        "tip_help_text": "",
                        "shortcut": ""
                    },
                    {
                        "text": "Paste Only Formula",
                        "command": ".uno:PasteOnlyFormula",
                        "style": 0,
                        "checked": false,
                        "enabled": true,
                        "default": false,
                        "help_command": "",
                        "help_text": "",
                        "tip_help_text": "",
                        "shortcut": ""
                    },
                    {
                        "text": "-"
                    },
                    {
                        "text": "Paste Transposed",
                        "command": ".uno:PasteTransposed",
                        "style": 0,
                        "checked": false,
                        "enabled": true,
                        "default": false,
                        "help_command": "",
                        "help_text": "",
                        "tip_help_text": "",
                        "shortcut": ""
                    },
                    {
                        "text": "-"
                    },
                    {
                        "text": "Paste ~Special...",
                        "command": ".uno:PasteSpecial",
                        "style": 0,
                        "checked": false,
                        "enabled": true,
                        "default": false,
                        "help_command": "",
                        "help_text": "",
                        "tip_help_text": "",
                        "shortcut": "Shift+Ctrl+V"
                    }
                ]
            },
            {
                "text": "-"
            },
            {
                "text": "Data Select",
                "command": ".uno:DataSelect",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": ""
            },
            {
                "text": "Current Validation",
                "command": ".uno:CurrentValidation",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": ""
            },
            {
                "text": "Define Current Name",
                "command": ".uno:DefineCurrentName",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": ""
            },
            {
                "text": "-"
            },
            {
                "text": "Insert ~Cells...",
                "command": ".uno:InsertCell",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": ""
            },
            {
                "text": "Delete C~ells...",
                "command": ".uno:DeleteCell",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": ""
            },
            {
                "text": "Delete",
                "command": ".uno:Delete",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": ""
            },
            {
                "text": "Merge Cells",
                "command": ".uno:MergeCells",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": ""
            },
            {
                "text": "Split Cell",
                "command": ".uno:SplitCell",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": ""
            },
            {
                "text": "-"
            },
            {
                "text": "Format Paintbrush",
                "command": ".uno:FormatPaintbrush",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": ""
            },
            {
                "text": "Reset Attributes",
                "command": ".uno:ResetAttributes",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": ""
            },
            {
                "text": "Format Styles Menu",
                "command": ".uno:FormatStylesMenu",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": "",
                "submenu": [
                    {
                        "text": "Edit Style",
                        "command": ".uno:EditStyle",
                        "style": 0,
                        "checked": false,
                        "enabled": true,
                        "default": false,
                        "help_command": "",
                        "help_text": "",
                        "tip_help_text": "",
                        "shortcut": ""
                    },
                    {
                        "text": "-"
                    },
                    {
                        "text": "Default Cell Styles",
                        "command": ".uno:DefaultCellStylesmenu",
                        "style": 2,
                        "checked": false,
                        "enabled": true,
                        "default": false,
                        "help_command": "",
                        "help_text": "",
                        "tip_help_text": "",
                        "shortcut": ""
                    },
                    {
                        "text": "Accent1 Cell Styles",
                        "command": ".uno:Accent1CellStyles",
                        "style": 2,
                        "checked": false,
                        "enabled": true,
                        "default": false,
                        "help_command": "",
                        "help_text": "",
                        "tip_help_text": "",
                        "shortcut": ""
                    },
                    {
                        "text": "Accent2 Cell Styles",
                        "command": ".uno:",
                        "style": 2,
                        "checked": false,
                        "enabled": true,
                        "default": false,
                        "help_command": "",
                        "help_text": "",
                        "tip_help_text": "",
                        "shortcut": ""
                    },
                    {
                        "text": "Accent 3 Cell Styles",
                        "command": ".uno:Accent3CellStyles",
                        "style": 2,
                        "checked": false,
                        "enabled": true,
                        "default": false,
                        "help_command": "",
                        "help_text": "",
                        "tip_help_text": "",
                        "shortcut": ""
                    },
                    {
                        "text": "-"
                    },
                    {
                        "text": "Bad Cell Styles",
                        "command": ".uno:BadCellStyles",
                        "style": 2,
                        "checked": false,
                        "enabled": true,
                        "default": false,
                        "help_command": "",
                        "help_text": "",
                        "tip_help_text": "",
                        "shortcut": ""
                    },
                    {
                        "text": "Error Cell Styles",
                        "command": ".uno:ErrorCellStyles",
                        "style": 2,
                        "checked": false,
                        "enabled": true,
                        "default": false,
                        "help_command": "",
                        "help_text": "",
                        "tip_help_text": "",
                        "shortcut": ""
                    },
                    {
                        "text": "Good Cell Styles",
                        "command": ".uno:GoodCellStyles",
                        "style": 2,
                        "checked": false,
                        "enabled": true,
                        "default": false,
                        "help_command": "",
                        "help_text": "",
                        "tip_help_text": "",
                        "shortcut": ""
                    },
                    {
                        "text": "Neutral Cell Styles",
                        "command": ".uno:NeutralCellStyles",
                        "style": 2,
                        "checked": false,
                        "enabled": true,
                        "default": false,
                        "help_command": "",
                        "help_text": "",
                        "tip_help_text": "",
                        "shortcut": ""
                    },
                    {
                        "text": "Warning Cell Styles",
                        "command": ".uno:WarningCellStyles",
                        "style": 2,
                        "checked": false,
                        "enabled": true,
                        "default": false,
                        "help_command": "",
                        "help_text": "",
                        "tip_help_text": "",
                        "shortcut": ""
                    },
                    {
                        "text": "-"
                    },
                    {
                        "text": "Footnote Cell Styles",
                        "command": ".uno:FootnoteCellStyles",
                        "style": 2,
                        "checked": false,
                        "enabled": true,
                        "default": false,
                        "help_command": "",
                        "help_text": "",
                        "tip_help_text": "",
                        "shortcut": ""
                    },
                    {
                        "text": "Note Cell Styles",
                        "command": ".uno:NoteCellStyles",
                        "style": 2,
                        "checked": false,
                        "enabled": true,
                        "default": false,
                        "help_command": "",
                        "help_text": "",
                        "tip_help_text": "",
                        "shortcut": ""
                    }
                ]
            },
            {
                "text": "-"
            },
            {
                "text": "Insert Annotation",
                "command": ".uno:InsertAnnotation",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": ""
            },
            {
                "text": "Edit Annotation",
                "command": ".uno:EditAnnotation",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": ""
            },
            {
                "text": "Delete Note",
                "command": ".uno:DeleteNote",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": ""
            },
            {
                "text": "Show Note",
                "command": ".uno:ShowNote",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": ""
            },
            {
                "text": "Hide Note",
                "command": ".uno:HideNote",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": ""
            },
            {
                "text": "-"
            },
            {
                "text": "Format Sparkline",
                "command": ".uno:FormatSparklineMenu",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": ""
            },
            {
                "text": "-"
            },
            {
                "text": "Conditional Formatting...",
                "command": ".uno:CurrentConditionalFormatDialog",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": ""
            },
            {
                "text": "Current Conditional Format Manager Dialog ...",
                "command": ".uno:CurrentConditionalFormatManagerDialog",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": ""
            },
            {
                "text": "Format Cell Dialog ...",
                "command": ".uno:FormatCellDialog",
                "style": 0,
                "checked": false,
                "enabled": true,
                "default": false,
                "help_command": "",
                "help_text": "",
                "tip_help_text": "",
                "shortcut": ""
            }
        ]
    }