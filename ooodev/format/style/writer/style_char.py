from __future__ import annotations
from typing import Tuple

from ....events.args.key_val_cancel_args import KeyValCancelArgs
from ....meta.static_prop import static_prop
from ...kind.format_kind import FormatKind
from ...style_base import StyleBase
from .kind.style_char_kind import StyleCharKind


class StyleChar(StyleBase):
    """
    Style Characters. Manages Chacter styles for Writer.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together Border Table properties.

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    def __init__(self, name: StyleCharKind | str = "") -> None:
        if name == "":
            name = StyleChar.default.prop_name
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
            event_args.value = StyleChar.default.prop_name

    # region Style Properties
    @property
    def bullet_symbols(self) -> StyleChar:
        """Style Bullet Symbols"""
        return StyleChar(StyleCharKind.BULLET_SYMBOLS)

    @property
    def caption_characters(self) -> StyleChar:
        """Style Caption characters"""
        return StyleChar(StyleCharKind.CAPTION_CHARACTERS)

    @property
    def citation(self) -> StyleChar:
        """Style Citation"""
        return StyleChar(StyleCharKind.CITATION)

    @property
    def definition(self) -> StyleChar:
        """Style Definition"""
        return StyleChar(StyleCharKind.DEFINITION)

    @property
    def drop_caps(self) -> StyleChar:
        """Style Drop Caps"""
        return StyleChar(StyleCharKind.DROP_CAPS)

    @property
    def emphasis(self) -> StyleChar:
        """Style Emphasis"""
        return StyleChar(StyleCharKind.EMPHASIS)

    @property
    def endnote_symbol(self) -> StyleChar:
        """Style Endnote Symbol"""
        return StyleChar(StyleCharKind.ENDNOTE_SYMBOL)

    @property
    def endnote_anchor(self) -> StyleChar:
        """Style Endnote anchor"""
        return StyleChar(StyleCharKind.ENDNOTE_ANCHOR)

    @property
    def example(self) -> StyleChar:
        """Style Example"""
        return StyleChar(StyleCharKind.EXAMPLE)

    @property
    def footnote_symbol(self) -> StyleChar:
        """Style Footnote Symbol"""
        return StyleChar(StyleCharKind.FOOTNOTE_SYMBOL)

    @property
    def footnote_anchor(self) -> StyleChar:
        """Style Footnote anchor"""
        return StyleChar(StyleCharKind.FOOTNOTE_ANCHOR)

    @property
    def index_link(self) -> StyleChar:
        """Style Index Link"""
        return StyleChar(StyleCharKind.INDEX_LINK)

    @property
    def internet_link(self) -> StyleChar:
        """Style Internet link"""
        return StyleChar(StyleCharKind.INTERNET_LINK)

    @property
    def line_numbering(self) -> StyleChar:
        """Style Line numbering"""
        return StyleChar(StyleCharKind.LINE_NUMBERING)

    @property
    def main_index_entry(self) -> StyleChar:
        """Style Main index entry"""
        return StyleChar(StyleCharKind.MAIN_INDEX_ENTRY)

    @property
    def numbering_symbols(self) -> StyleChar:
        """Style Numbering Symbols"""
        return StyleChar(StyleCharKind.NUMBERING_SYMBOLS)

    @property
    def page_number(self) -> StyleChar:
        """Style Page Number"""
        return StyleChar(StyleCharKind.PAGE_NUMBER)

    @property
    def placeholder(self) -> StyleChar:
        """Style Placeholder"""
        return StyleChar(StyleCharKind.PLACEHOLDER)

    @property
    def rubies(self) -> StyleChar:
        """Style Rubies"""
        return StyleChar(StyleCharKind.RUBIES)

    @property
    def source_text(self) -> StyleChar:
        """Style Source Text"""
        return StyleChar(StyleCharKind.SOURCE_TEXT)

    @property
    def standard(self) -> StyleChar:
        """Removes styling"""
        return StyleChar(StyleCharKind.STANDARD)

    @property
    def strong_emphasis(self) -> StyleChar:
        """Style Strong Emphasis"""
        return StyleChar(StyleCharKind.STRONG_EMPHASIS)

    @property
    def teletype(self) -> StyleChar:
        """Style Teletype"""
        return StyleChar(StyleCharKind.TELETYPE)

    @property
    def user_entry(self) -> StyleChar:
        """Style User Entry"""
        return StyleChar(StyleCharKind.USER_ENTRY)

    @property
    def variable(self) -> StyleChar:
        """Style Variable"""
        return StyleChar(StyleCharKind.VARIABLE)

    @property
    def vertical_numbering_symbols(self) -> StyleChar:
        """Style Vertical Numbering Symbols"""
        return StyleChar(StyleCharKind.VERTICAL_NUMBERING_SYMBOLS)

    @property
    def visited_internet_link(self) -> StyleChar:
        """Style Visited Internet Link"""
        return StyleChar(StyleCharKind.VISITED_INTERNET_LINK)

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
        if self is StyleChar.default:
            raise ValueError("Setting StyleChar.default properties is not allowed.")
        self._set(self._get_property_name(), str(value))

    @static_prop
    def default() -> StyleChar:  # type: ignore[misc]
        """Gets ``StyleChar`` default. Static Property."""
        if StyleChar._DEFAULT is None:
            StyleChar._DEFAULT = StyleChar(name="Standard")
        return StyleChar._DEFAULT
