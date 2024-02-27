from __future__ import annotations
import contextlib
from enum import Enum
from ooodev.utils.kind import kind_helper


class DialogControlKind(Enum):
    """
    Dialog Control Values.
    """

    BUTTON = "com.sun.star.awt.UnoControlButton"
    CHECKBOX = "com.sun.star.awt.UnoControlCheckBox"
    COMBOBOX = "com.sun.star.awt.UnoControlComboBox"
    CURRENCY = "com.sun.star.awt.UnoControlCurrencyField"
    DATE_FIELD = "com.sun.star.awt.UnoControlDateField"
    FILE_CONTROL = "com.sun.star.awt.UnoControlFileControl"
    FIXED_LINE = "com.sun.star.awt.UnoControlFixedLine"
    FIXED_TEXT = "com.sun.star.awt.UnoControlFixedText"
    FORMATTED_TEXT = "com.sun.star.awt.UnoControlFormattedField"
    GRID_CONTROL = "com.sun.star.awt.grid.UnoControlGrid"
    GROUP_BOX = "com.sun.star.awt.UnoControlGroupBox"
    HYPERLINK = "com.sun.star.awt.UnoControlFixedHyperlink"
    IMAGE = "com.sun.star.awt.UnoControlImageControl"
    LIST_BOX = "com.sun.star.awt.UnoControlListBox"
    NUMERIC = "com.sun.star.awt.UnoControlNumericField"
    PATTERN = "com.sun.star.awt.UnoControlPatternField"
    PROGRESS_BAR = "com.sun.star.awt.UnoControlProgressBar"
    RADIO_BUTTON = "com.sun.star.awt.UnoControlRadioButton"
    SCROLL_BAR = "com.sun.star.awt.UnoControlScrollBar"
    SPIN_BUTTON = "com.sun.star.awt.UnoControlSpinButton"
    TAB_PAGE_CONTAINER = "com.sun.star.awt.tab.UnoControlTabPageContainer"
    TAB_PAGE = "com.sun.star.awt.tab.UnoControlTabPage"
    EDIT = "com.sun.star.awt.UnoControlEdit"
    TIME = "com.sun.star.awt.UnoControlTimeField"
    TREE = "com.sun.star.awt.tree.TreeControl"
    UNKNOWN = "UNKNOWN"

    def __str__(self) -> str:
        return self.value

    @staticmethod
    def from_str(s: str) -> "DialogControlKind":
        """
        Gets an ``DialogControlKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Returns:
            DialogControlKind: Enum instance.
        """
        with contextlib.suppress(Exception):
            return kind_helper.enum_from_string(s, DialogControlKind)
        return DialogControlKind.UNKNOWN

    @staticmethod
    def from_value(s: str) -> "DialogControlKind":
        """
        Gets an ``DialogControlKind`` instance from string.

        Args:
            s (str): String that represents the value of an enums values.

        Returns:
            DialogControlKind: Enum instance.
        """
        with contextlib.suppress(Exception):
            return DialogControlKind(s)
        return DialogControlKind.UNKNOWN
