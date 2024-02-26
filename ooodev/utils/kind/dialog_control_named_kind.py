from __future__ import annotations
import contextlib
from enum import Enum
from ooodev.utils.kind import kind_helper


class DialogControlNamedKind(Enum):
    """
    Dialog Control Values.
    """

    BUTTON = "stardiv.Toolkit.UnoButtonControl"
    CHECKBOX = "stardiv.Toolkit.UnoCheckBoxControl"
    COMBOBOX = "stardiv.Toolkit.UnoComboBoxControl"
    CURRENCY = "stardiv.Toolkit.UnoCurrencyFieldControl"
    DATE_FIELD = "stardiv.Toolkit.UnoDateFieldControl"
    FILE_CONTROL = "stardiv.Toolkit.UnoFileControl"
    FIXED_LINE = "stardiv.Toolkit.UnoFixedLineControl"
    FIXED_TEXT = "stardiv.Toolkit.UnoFixedTextControl"
    FORMATTED_TEXT = "stardiv.Toolkit.UnoFormattedFieldControl"
    GRID_CONTROL = "stardiv.Toolkit.GridControl"
    GROUP_BOX = "stardiv.Toolkit.UnoGroupBoxControl"
    HYPERLINK = "stardiv.Toolkit.UnoFixedHyperlinkControl"
    IMAGE = "stardiv.Toolkit.UnoImageControlControl"
    LIST_BOX = "stardiv.Toolkit.UnoListBoxControl"
    NUMERIC = "stardiv.Toolkit.UnoNumericFieldControl"
    PATTERN = "stardiv.Toolkit.UnoPatternFieldControl"
    PROGRESS_BAR = "stardiv.Toolkit.UnoProgressBarControl"
    RADIO_BUTTON = "stardiv.Toolkit.UnoRadioButtonControl"
    SCROLL_BAR = "stardiv.Toolkit.UnoScrollBarControl"
    SPIN_BUTTON = "stardiv.Toolkit.UnoSpinButtonControl"
    TAB_PAGE_CONTAINER = "stardiv.Toolkit.UnoControlTabPageContainer"
    TAB_PAGE = "stardiv.Toolkit.UnoControlTabPage"
    EDIT = "stardiv.Toolkit.UnoEditControl"
    TIME = "stardiv.Toolkit.UnoTimeFieldControl"
    TREE = "stardiv.Toolkit.UnoTreeControl"
    UNKNOWN = "UNKNOWN"

    def __str__(self) -> str:
        return self.value

    @staticmethod
    def from_str(s: str) -> "DialogControlNamedKind":
        """
        Gets an ``DialogControlNamedKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Returns:
            DialogControlNamedKind: Enum instance.
        """
        with contextlib.suppress(Exception):
            return kind_helper.enum_from_string(s, DialogControlNamedKind)
        return DialogControlNamedKind.UNKNOWN

    @staticmethod
    def from_value(s: str) -> "DialogControlNamedKind":
        """
        Gets an ``DialogControlNamedKind`` instance from string.

        Args:
            s (str): String that represents the value of an enums values.

        Returns:
            DialogControlNamedKind: Enum instance.
        """
        with contextlib.suppress(Exception):
            return DialogControlNamedKind(s)
        return DialogControlNamedKind.UNKNOWN
