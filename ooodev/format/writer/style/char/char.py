# region Import
from __future__ import annotations
from typing import Any, Tuple

from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.meta.static_prop import static_prop
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleName
from ooodev.format.writer.style.char.kind.style_char_kind import StyleCharKind as StyleCharKind

# endregion Import


class Char(StyleName):
    """
    Characters Style.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together Border Table properties.

    .. seealso::

        - :ref:`help_writer_format_style_char`

    .. versionadded:: 0.9.0
    """

    def __init__(self, name: StyleCharKind | str = "") -> None:
        """
        Constructor

        Args:
            name (StyleCharKind, str, optional): Style Name. Defaults to "Standard".

        Returns:
            None:

        See Also:

            - :ref:`help_writer_format_style_char`
        """
        if name == "":
            name = Char.default.prop_name
        super().__init__(name=name)

    # region Overrides

    def _get_family_style_name(self) -> str:
        return "CharacterStyles"

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.CharacterProperties",)
        return self._supported_services_values

    def _get_property_name(self) -> str:
        try:
            return self._style_property_name
        except AttributeError:
            self._style_property_name = "CharStyleName"
        return self._style_property_name

    def on_property_setting(self, source: Any, event_args: KeyValCancelArgs):
        """
        Triggers for each property that is set

        Args:
            event_args (KeyValueCancelArgs): Event Args
        """
        # there is only one style property for this class.
        # if CharStyleName is set to "" then an error is raised.
        # Solution is set to "No Character Style" or "Standard" Which LibreOffice recognizes and set to ""
        # this event covers apply() and restore()
        if event_args.value == "":
            event_args.value = Char.default.prop_name
        super().on_property_setting(source, event_args)

    # endregion Overrides

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

    # region Properties

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.STYLE | FormatKind.CHAR
        return self._format_kind_prop

    @static_prop
    def default() -> Char:  # type: ignore[misc]
        """Gets ``StyleChar`` default. Static Property."""
        try:
            return Char._DEFAULT_CHAR  # type: ignore[attr-defined]
        except AttributeError:
            Char._DEFAULT_CHAR = Char(name="Standard")  # type: ignore[attr-defined]
            Char._DEFAULT_CHAR._is_default_inst = True  # type: ignore[attr-defined]
        return Char._DEFAULT_CHAR  # type: ignore[attr-defined]

    # endregion Properties
