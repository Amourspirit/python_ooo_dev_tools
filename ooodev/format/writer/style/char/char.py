from __future__ import annotations
from typing import Tuple

from .....events.args.key_val_cancel_args import KeyValCancelArgs
from .....meta.static_prop import static_prop
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase
from .kind import StyleCharKind as StyleCharKind


class Char(StyleBase):
    """
    Style Characters. Manages Chacter styles for Writer.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together Border Table properties.

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    def __init__(self, name: StyleCharKind | str = "") -> None:
        if name == "":
            name = Char.default.prop_name
        super().__init__(**{self._get_property_name(): str(name)})

    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.style.CharacterProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.CharacterProperties",)

    def _get_property_name(self) -> str:
        return "CharStyleName"

    def on_property_setting(self, event_args: KeyValCancelArgs):
        """
        Triggers for each property that is set

        Args:
            event_args (KeyValueCancelArgs): Event Args
        """
        # there is only one style property for this class.
        # if CharStyleName is set to "" then an error is raised.
        # Solution is set to "No Character Style" or "Standard" Which LibreOffice recognizes and set to ""
        # this event covers apply() and resore()
        if event_args.value == "":
            event_args.value = Char.default.prop_name

    # region Style Properties
    @property
    def bullet_symbols(self) -> Char:
        """Style Bullet Symbols"""
        return Char(StyleCharKind.BULLET_SYMBOLS)

    @property
    def caption_characters(self) -> Char:
        """Style Caption characters"""
        return Char(StyleCharKind.CAPTION_CHARACTERS)

    @property
    def citation(self) -> Char:
        """Style Citation"""
        return Char(StyleCharKind.CITATION)

    @property
    def quotation(self) -> Char:
        """Style Citation"""
        return Char(StyleCharKind.CITATION)

    @property
    def definition(self) -> Char:
        """Style Definition"""
        return Char(StyleCharKind.DEFINITION)

    @property
    def drop_caps(self) -> Char:
        """Style Drop Caps"""
        return Char(StyleCharKind.DROP_CAPS)

    @property
    def emphasis(self) -> Char:
        """Style Emphasis"""
        return Char(StyleCharKind.EMPHASIS)

    @property
    def endnote_symbol(self) -> Char:
        """Style Endnote Symbol"""
        return Char(StyleCharKind.ENDNOTE_SYMBOL)

    @property
    def endnote_anchor(self) -> Char:
        """Style Endnote anchor"""
        return Char(StyleCharKind.ENDNOTE_ANCHOR)

    @property
    def example(self) -> Char:
        """Style Example"""
        return Char(StyleCharKind.EXAMPLE)

    @property
    def footnote_symbol(self) -> Char:
        """Style Footnote Symbol"""
        return Char(StyleCharKind.FOOTNOTE_SYMBOL)

    @property
    def footnote_anchor(self) -> Char:
        """Style Footnote anchor"""
        return Char(StyleCharKind.FOOTNOTE_ANCHOR)

    @property
    def index_link(self) -> Char:
        """Style Index Link"""
        return Char(StyleCharKind.INDEX_LINK)

    @property
    def internet_link(self) -> Char:
        """Style Internet link"""
        return Char(StyleCharKind.INTERNET_LINK)

    @property
    def line_numbering(self) -> Char:
        """Style Line numbering"""
        return Char(StyleCharKind.LINE_NUMBERING)

    @property
    def main_index_entry(self) -> Char:
        """Style Main index entry"""
        return Char(StyleCharKind.MAIN_INDEX_ENTRY)

    @property
    def numbering_symbols(self) -> Char:
        """Style Numbering Symbols"""
        return Char(StyleCharKind.NUMBERING_SYMBOLS)

    @property
    def page_number(self) -> Char:
        """Style Page Number"""
        return Char(StyleCharKind.PAGE_NUMBER)

    @property
    def placeholder(self) -> Char:
        """Style Placeholder"""
        return Char(StyleCharKind.PLACEHOLDER)

    @property
    def rubies(self) -> Char:
        """Style Rubies"""
        return Char(StyleCharKind.RUBIES)

    @property
    def source_text(self) -> Char:
        """Style Source Text"""
        return Char(StyleCharKind.SOURCE_TEXT)

    @property
    def standard(self) -> Char:
        """Removes styling"""
        return Char(StyleCharKind.STANDARD)

    @property
    def strong_emphasis(self) -> Char:
        """Style Strong Emphasis"""
        return Char(StyleCharKind.STRONG_EMPHASIS)

    @property
    def teletype(self) -> Char:
        """Style Teletype"""
        return Char(StyleCharKind.TELETYPE)

    @property
    def user_entry(self) -> Char:
        """Style User Entry"""
        return Char(StyleCharKind.USER_ENTRY)

    @property
    def variable(self) -> Char:
        """Style Variable"""
        return Char(StyleCharKind.VARIABLE)

    @property
    def vertical_numbering_symbols(self) -> Char:
        """Style Vertical Numbering Symbols"""
        return Char(StyleCharKind.VERTICAL_NUMBERING_SYMBOLS)

    @property
    def visited_internet_link(self) -> Char:
        """Style Visited Internet Link"""
        return Char(StyleCharKind.VISITED_INTERNET_LINK)

    # endregion Style Properties

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.STYLE | FormatKind.CHAR

    @property
    def prop_name(self) -> str:
        """Gets/Sets Character style namd"""
        return self._get(self._get_property_name())

    @prop_name.setter
    def prop_name(self, value: StyleCharKind | str) -> None:
        if self is Char.default:
            raise ValueError("Setting StyleChar.default properties is not allowed.")
        self._set(self._get_property_name(), str(value))

    @static_prop
    def default() -> Char:  # type: ignore[misc]
        """Gets ``StyleChar`` default. Static Property."""
        if Char._DEFAULT is None:
            Char._DEFAULT = Char(name="Standard")
        return Char._DEFAULT
