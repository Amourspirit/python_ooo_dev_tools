from __future__ import annotations
from typing import Tuple

from ....events.args.key_val_cancel_args import KeyValCancelArgs
from ....meta.static_prop import static_prop
from ...kind.format_kind import FormatKind
from ...style_base import StyleBase
from .kind.style_para_kind import StyleParaKind as StyleParaKind


class StylePara(StyleBase):
    """
    Style Paragraph. Manages Paragraph styles for Writer.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together Border Table properties.

    .. versionadded:: 0.9.0
    """
    _DEFAULT = None

    def __init__(self, name: StyleParaKind | str = "") -> None:
        if name == "":
            name = StylePara.default.prop_name
        super().__init__(**{self._get_property_name(): str(name)})

    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.style.ParagraphProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.ParagraphProperties",)

    def _get_property_name(self) -> str:
        return "ParaStyleName"

    def on_property_setting(self, event_args: KeyValCancelArgs):
        """
        Triggers for each property that is set

        Args:
            event_args (KeyValueCancelArgs): Event Args
        """
        # there is only one style property for this class.
        # if ParaStyleName is set to "" then an error is raised.
        # Solution is set to "Standard" Which LibreOffice recognizes and set to ""
        # this event covers apply() and resore()
        if event_args.value == "":
            event_args.value = StylePara.default.prop_name

    # region Style Properties
    @property
    def addressee(self) -> StylePara:
        """Style Addressee"""
        return StylePara(StyleParaKind.ADDRESSEE)

    @property
    def addressee(self) -> StylePara:
        """Style Addressee"""
        return StylePara(StyleParaKind.ADDRESSEE)

    @property
    def salutation(self) -> StylePara:
        """Style Salutation"""
        return StylePara(StyleParaKind.SALUTATION)

    @property
    def complinentary_close(self) -> StylePara:
        """Style Complinentary Close"""
        return StylePara(StyleParaKind.SALUTATION)

    @property
    def endnote(self) -> StylePara:
        """Style Endnote"""
        return StylePara(StyleParaKind.ENDNOTE)

    @property
    def footnote(self) -> StylePara:
        """Style Footnote"""
        return StylePara(StyleParaKind.FOOTNOTE)

    @property
    def frame_contents(self) -> StylePara:
        """Style Frame Contents"""
        return StylePara(StyleParaKind.FRAME_CONTENTS)

    @property
    def header(self) -> StylePara:
        """Style Header"""
        return StylePara(StyleParaKind.HEADER)

    @property
    def header_footer(self) -> StylePara:
        """Style Header Footer"""
        return StylePara(StyleParaKind.HEADER_FOOTER)

    @property
    def header_left(self) -> StylePara:
        """Style Header Left"""
        return StylePara(StyleParaKind.HEADER_LEFT)

    @property
    def header_right(self) -> StylePara:
        """Style Header Right"""
        return StylePara(StyleParaKind.HEADER_RIGHT)

    @property
    def footer(self) -> StylePara:
        """Style Footer"""
        return StylePara(StyleParaKind.FOOTER)

    @property
    def footer_left(self) -> StylePara:
        """Style Footer Left"""
        return StylePara(StyleParaKind.FOOTER_LEFT)

    @property
    def footer_right(self) -> StylePara:
        """Style Footer Right"""
        return StylePara(StyleParaKind.FOOTER_RIGHT)

    @property
    def heading(self) -> StylePara:
        """Style Heading"""
        return StylePara(StyleParaKind.HEADING)

    @property
    def h1(self) -> StylePara:
        """Style Heading 1"""
        return StylePara(StyleParaKind.HEADING_1)

    @property
    def h2(self) -> StylePara:
        """Style Heading 2"""
        return StylePara(StyleParaKind.HEADING_2)

    @property
    def h3(self) -> StylePara:
        """Style Heading 3"""
        return StylePara(StyleParaKind.HEADING_3)

    @property
    def h4(self) -> StylePara:
        """Style Heading 4"""
        return StylePara(StyleParaKind.HEADING_4)

    @property
    def h5(self) -> StylePara:
        """Style Heading 5"""
        return StylePara(StyleParaKind.HEADING_5)

    @property
    def h6(self) -> StylePara:
        """Style Heading 6"""
        return StylePara(StyleParaKind.HEADING_6)

    @property
    def h7(self) -> StylePara:
        """Style Heading 7"""
        return StylePara(StyleParaKind.HEADING_7)

    @property
    def h8(self) -> StylePara:
        """Style Heading 8"""
        return StylePara(StyleParaKind.HEADING_8)

    @property
    def h9(self) -> StylePara:
        """Style Heading 9"""
        return StylePara(StyleParaKind.HEADING_9)

    @property
    def h10(self) -> StylePara:
        """Style Heading 10"""
        return StylePara(StyleParaKind.HEADING_10)

    @property
    def horizontal_line(self) -> StylePara:
        """Style Horizontal Line"""
        return StylePara(StyleParaKind.HORIZONTAL_LINE)

    @property
    def idx(self) -> StylePara:
        """Style Index"""
        return StylePara(StyleParaKind.INDEX)

    @property
    def idx_bib1(self) -> StylePara:
        """Style Bibliography 1"""
        return StylePara(StyleParaKind.BIBLIOGRAPHY_1)

    @property
    def idx_c1(self) -> StylePara:
        """Style Contents 1"""
        return StylePara(StyleParaKind.CONTENTS_1)

    @property
    def idx_c2(self) -> StylePara:
        """Style Contents 2"""
        return StylePara(StyleParaKind.CONTENTS_2)

    @property
    def idx_c3(self) -> StylePara:
        """Style Contents 3"""
        return StylePara(StyleParaKind.CONTENTS_3)

    @property
    def idx_c4(self) -> StylePara:
        """Style Contents 4"""
        return StylePara(StyleParaKind.CONTENTS_4)

    @property
    def idx_c5(self) -> StylePara:
        """Style Contents 5"""
        return StylePara(StyleParaKind.CONTENTS_5)

    @property
    def idx_c6(self) -> StylePara:
        """Style Contents 6"""
        return StylePara(StyleParaKind.CONTENTS_6)

    @property
    def idx_c7(self) -> StylePara:
        """Style Contents 7"""
        return StylePara(StyleParaKind.CONTENTS_7)

    @property
    def idx_c8(self) -> StylePara:
        """Style Contents 8"""
        return StylePara(StyleParaKind.CONTENTS_8)

    @property
    def idx_c9(self) -> StylePara:
        """Style Contents 9"""
        return StylePara(StyleParaKind.CONTENTS_9)

    @property
    def idx_c10(self) -> StylePara:
        """Style Contents 10"""
        return StylePara(StyleParaKind.CONTENTS_10)

    @property
    def idx_separator(self) -> StylePara:
        """Style Index Separator"""
        return StylePara(StyleParaKind.INDEX_SEPARATOR)

    @property
    def idx_1(self) -> StylePara:
        """Style Index 1"""
        return StylePara(StyleParaKind.INDEX_1)

    @property
    def idx_2(self) -> StylePara:
        """Style Index 2"""
        return StylePara(StyleParaKind.INDEX_2)

    @property
    def idx_3(self) -> StylePara:
        """Style Index 3"""
        return StylePara(StyleParaKind.INDEX_3)

    @property
    def idx_obj1(self) -> StylePara:
        """Style Object Index 1"""
        return StylePara(StyleParaKind.OBJECT_INDEX_1)

    @property
    def idx_tbl1(self) -> StylePara:
        """Style Table Index 1"""
        return StylePara(StyleParaKind.TABLE_INDEX_1)

    @property
    def idx_user1(self) -> StylePara:
        """Style User Index 1"""
        return StylePara(StyleParaKind.USER_INDEX_1)

    @property
    def idx_user2(self) -> StylePara:
        """Style User Index 2"""
        return StylePara(StyleParaKind.USER_INDEX_2)

    @property
    def idx_user3(self) -> StylePara:
        """Style User Index 3"""
        return StylePara(StyleParaKind.USER_INDEX_3)

    @property
    def idx_user4(self) -> StylePara:
        """Style User Index 4"""
        return StylePara(StyleParaKind.USER_INDEX_4)

    @property
    def idx_user5(self) -> StylePara:
        """Style User Index 5"""
        return StylePara(StyleParaKind.USER_INDEX_5)

    @property
    def idx_user6(self) -> StylePara:
        """Style User Index 6"""
        return StylePara(StyleParaKind.USER_INDEX_6)

    @property
    def idx_user7(self) -> StylePara:
        """Style User Index 7"""
        return StylePara(StyleParaKind.USER_INDEX_7)

    @property
    def idx_user8(self) -> StylePara:
        """Style User Index 8"""
        return StylePara(StyleParaKind.USER_INDEX_8)

    @property
    def idx_user9(self) -> StylePara:
        """Style User Index 9"""
        return StylePara(StyleParaKind.USER_INDEX_9)

    @property
    def idx_user10(self) -> StylePara:
        """Style User Index 10"""
        return StylePara(StyleParaKind.USER_INDEX_10)

    @property
    def list_contents(self) -> StylePara:
        """Style List Contents"""
        return StylePara(StyleParaKind.LIST_CONTENTS)

    @property
    def list_heading(self) -> StylePara:
        """Style List Heading"""
        return StylePara(StyleParaKind.LIST_HEADING)

    @property
    def pre_text(self) -> StylePara:
        """Style Preformatted text"""
        return StylePara(StyleParaKind.PREFORMATTED_TEXT)

    @property
    def quotations(self) -> StylePara:
        """Style Quotations"""
        return StylePara(StyleParaKind.QUOTATIONS)

    @property
    def sender(self) -> StylePara:
        """Style Sender"""
        return StylePara(StyleParaKind.SENDER)

    @property
    def signature(self) -> StylePara:
        """Style Signature"""
        return StylePara(StyleParaKind.SIGNATURE)

    @property
    def tbl_contents(self) -> StylePara:
        """Style Table Contents"""
        return StylePara(StyleParaKind.TABLE_CONTENTS)

    @property
    def tbl_heading(self) -> StylePara:
        """Style Table Heading"""
        return StylePara(StyleParaKind.TABLE_HEADING)

    @property
    def txt_body(self) -> StylePara:
        """Style Text Body"""
        return StylePara(StyleParaKind.TEXT_BODY)

    @property
    def txt_first_line_indent(self) -> StylePara:
        """Style First Line Indent"""
        return StylePara(StyleParaKind.FIRST_LINE_INDENT)

    @property
    def txt_hanging_indent(self) -> StylePara:
        """Style Hanging Indent"""
        return StylePara(StyleParaKind.HANGING_INDENT)

    @property
    def txt_list_indent(self) -> StylePara:
        """Style List Indent"""
        return StylePara(StyleParaKind.LIST_INDENT)

    @property
    def txt_marginalia(self) -> StylePara:
        """Style Marginalia"""
        return StylePara(StyleParaKind.MARGINALIA)

    @property
    def txt_body_indent(self) -> StylePara:
        """Style Text Body Indent"""
        return StylePara(StyleParaKind.TEXT_BODY_INDENT)

    # endregion Style Properties

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.STYLE | FormatKind.PARA

    @property
    def prop_name(self) -> str:
        """Gets/Sets Character style namd"""
        return self._get(self._get_property_name())

    @prop_name.setter
    def prop_name(self, value: StyleParaKind | str) -> None:
        if self is StylePara.default:
            raise ValueError("Setting StylePara.default properties is not allowed.")
        self._set(self._get_property_name(), str(value))

    @static_prop
    def default() -> StylePara:  # type: ignore[misc]
        """Gets ``StylePara`` default. Static Property."""
        if StylePara._DEFAULT is None:
            StylePara._DEFAULT = StylePara(name=StyleParaKind.STANDARD)
        return StylePara._DEFAULT
