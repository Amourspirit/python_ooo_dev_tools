from .kind_base import KindBase
from ooodev.utils.kind import kind_helper


class FormControlKind(KindBase):
    """
    Values used with ``com.sun.star.form.control.*``

    See Also:
        `Control API <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1form_1_1control.html>`_
    """

    CHECK_BOX = "CheckBox"
    COMBO_BOX = "ComboBox"
    COMMAND_BUTTON = "CommandButton"
    CURRENCY_FIELD = "CurrencyField"
    DATE_FIELD = "DateField"
    FILTER_CONTROL = "FilterControl"
    FORMATTED_FIELD = "FormattedField"
    GRID_CONTROL = "GridControl"
    GROUP_BOX = "GroupBox"
    IMAGE_BUTTON = "ImageButton"
    IMAGE_CONTROL = "ImageControl"
    INTERACTION_GRID_CONTROL = "InteractionGridControl"
    LIST_BOX = "ListBox"
    NAVIGATION_TOOL_BAR = "NavigationToolBar"
    NUMERIC_FIELD = "NumericField"
    PATTERN_FIELD = "PatternField"
    RADIO_BUTTON = "RadioButton"
    SUBMIT_BUTTON = "SubmitButton"
    TEXT_FIELD = "TextField"
    TIME_FIELD = "TimeField"

    def to_namespace(self) -> str:
        """Gets full name-space value of instance"""
        return f"com.sun.star.form.control.{self.value}"

    @staticmethod
    def from_str(s: str) -> "FormControlKind":
        """
        Gets an ``FormControlKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hyphen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``FormControlKind`` instance.

        Returns:
            FormControlKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, FormControlKind)
