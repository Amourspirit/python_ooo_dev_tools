from .kind_base import KindBase
from . import kind_helper


class FormComponentKind(KindBase):
    """
    Values used with ``com.sun.star.form.component.*``

    See Also:
        `component API <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1form_1_1component.html>`_
    """

    CHECK_BOX = "CheckBox"
    COMBO_BOX = "ComboBox"
    COMMAND_BUTTON = "CommandButton"
    CURRENCY_FIELD = "CurrencyField"
    DATABASE_CHECK_BOX = "DatabaseCheckBox"
    DATABASE_COMBO_BOX = "DatabaseComboBox"
    DATABASE_CURRENCY_FIELD = "DatabaseCurrencyField"
    DATABASE_DATE_FIELD = "DatabaseDateField"
    DATABASE_FORMATTED_FIELD = "DatabaseFormattedField"
    DATABASE_IMAGE_CONTROL = "DatabaseImageControl"
    DATABASE_LIST_BOX = "DatabaseListBox"
    DATABASE_NUMERIC_FIELD = "DatabaseNumericField"
    DATABASE_PATTERN_FIELD = "DatabasePatternField"
    DATABASE_RADIO_BUTTON = "DatabaseRadioButton"
    DATABASE_TEXT_FIELD = "DatabaseTextField"
    DATABASE_TIME_FIELD = "DatabaseTimeField"
    DATE_FIELD = "DateField"
    FILE_CONTROL = "FileControl"
    FIXED_TEXT = "FixedText"
    FORMATTED_FIELD = "FormattedField"
    GRID_CONTROL = "GridControl"
    GROUP_BOX = "GroupBox"
    HIDDEN_CONTROL = "HiddenControl"
    HTML_FORM = "HTMLForm"
    IMAGE_BUTTON = "ImageButton"
    LIST_BOX = "ListBox"
    NAVIGATION_TOOL_BAR = "NavigationToolBar"
    NUMERIC_FIELD = "NumericField"
    PATTERN_FIELD = "PatternField"
    RADIO_BUTTON = "RadioButton"
    RICH_TEXT_CONTROL = "RichTextControl"
    SCROLL_BAR = "ScrollBar"
    SPIN_BUTTON = "SpinButton"
    SUBMIT_BUTTON = "SubmitButton"
    TEXT_FIELD = "TextField"
    TIME_FIELD = "TimeField"

    def to_namespace(self) -> str:
        """Gets full name-space value of instance"""
        return f"com.sun.star.form.component.{self.value}"

    @staticmethod
    def from_str(s: str) -> "FormComponentKind":
        """
        Gets an ``FormComponentKind`` instance from string.

        Args:
            s (str): String that represents the name of an enum Name.
                ``s`` is case insensitive and can be ``CamelCase``, ``pascal_case`` , ``snake_case``,
                ``hypen-case``, ``normal case``.

        Raises:
            ValueError: If input string is empty.
            AttributeError: If unable to get ``FormComponentKind`` instance.

        Returns:
            FormComponentKind: Enum instance.
        """
        return kind_helper.enum_from_string(s, FormComponentKind)
