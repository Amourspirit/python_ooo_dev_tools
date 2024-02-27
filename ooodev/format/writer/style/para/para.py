# region Import
from __future__ import annotations
from typing import Any, Tuple

from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.meta.static_prop import static_prop
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleName
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind as StyleParaKind

# endregion Import


class Para(StyleName):
    """
    Style Paragraph. Manages Paragraph styles for Writer.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together Border Table properties.

    .. seealso::

        - :ref:`help_writer_format_style_para`

    .. versionadded:: 0.9.0
    """

    def __init__(self, name: StyleParaKind | str = "") -> None:
        """
        Constructor

        Args:
            name (StyleCharKind, str, optional): Style Name. Defaults to "Standard".

        Returns:
            None:

        See Also:

            - :ref:`help_writer_format_style_para`
        """
        if name == "":
            name = Para.default.prop_name
        super().__init__(name=name)

    # region Overrides
    def _get_family_style_name(self) -> str:
        return "ParagraphStyles"

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.ParagraphProperties",)
        return self._supported_services_values

    def _get_property_name(self) -> str:
        try:
            return self._property_name
        except AttributeError:
            self._property_name = "ParaStyleName"
        return self._property_name

    def on_property_setting(self, source: Any, event_args: KeyValCancelArgs):
        """
        Triggers for each property that is set

        Args:
            event_args (KeyValueCancelArgs): Event Args
        """
        # there is only one style property for this class.
        # if ParaStyleName is set to "" then an error is raised.
        # Solution is set to "Standard" Which LibreOffice recognizes and set to ""
        # this event covers apply() and restore()
        if event_args.value == "":
            event_args.value = Para.default.prop_name
        super().on_property_setting(source, event_args)

    # endregion Overrides

    # region Style Properties
    @property
    def addressee(self) -> Para:
        """Style Addressee"""
        return Para(StyleParaKind.ADDRESSEE)

    @property
    def salutation(self) -> Para:
        """Style Salutation"""
        return Para(StyleParaKind.SALUTATION)

    @property
    def complimentary_close(self) -> Para:
        """Style Complimentary Close"""
        return Para(StyleParaKind.SALUTATION)

    @property
    def endnote(self) -> Para:
        """Style Endnote"""
        return Para(StyleParaKind.ENDNOTE)

    @property
    def footnote(self) -> Para:
        """Style Footnote"""
        return Para(StyleParaKind.FOOTNOTE)

    @property
    def frame_contents(self) -> Para:
        """Style Frame Contents"""
        return Para(StyleParaKind.FRAME_CONTENTS)

    @property
    def header(self) -> Para:
        """Style Header"""
        return Para(StyleParaKind.HEADER)

    @property
    def header_footer(self) -> Para:
        """Style Header Footer"""
        return Para(StyleParaKind.HEADER_FOOTER)

    @property
    def header_left(self) -> Para:
        """Style Header Left"""
        return Para(StyleParaKind.HEADER_LEFT)

    @property
    def header_right(self) -> Para:
        """Style Header Right"""
        return Para(StyleParaKind.HEADER_RIGHT)

    @property
    def footer(self) -> Para:
        """Style Footer"""
        return Para(StyleParaKind.FOOTER)

    @property
    def footer_left(self) -> Para:
        """Style Footer Left"""
        return Para(StyleParaKind.FOOTER_LEFT)

    @property
    def footer_right(self) -> Para:
        """Style Footer Right"""
        return Para(StyleParaKind.FOOTER_RIGHT)

    @property
    def heading(self) -> Para:
        """Style Heading"""
        return Para(StyleParaKind.HEADING)

    @property
    def h1(self) -> Para:
        """Style Heading 1"""
        return Para(StyleParaKind.HEADING_1)

    @property
    def h2(self) -> Para:
        """Style Heading 2"""
        return Para(StyleParaKind.HEADING_2)

    @property
    def h3(self) -> Para:
        """Style Heading 3"""
        return Para(StyleParaKind.HEADING_3)

    @property
    def h4(self) -> Para:
        """Style Heading 4"""
        return Para(StyleParaKind.HEADING_4)

    @property
    def h5(self) -> Para:
        """Style Heading 5"""
        return Para(StyleParaKind.HEADING_5)

    @property
    def h6(self) -> Para:
        """Style Heading 6"""
        return Para(StyleParaKind.HEADING_6)

    @property
    def h7(self) -> Para:
        """Style Heading 7"""
        return Para(StyleParaKind.HEADING_7)

    @property
    def h8(self) -> Para:
        """Style Heading 8"""
        return Para(StyleParaKind.HEADING_8)

    @property
    def h9(self) -> Para:
        """Style Heading 9"""
        return Para(StyleParaKind.HEADING_9)

    @property
    def h10(self) -> Para:
        """Style Heading 10"""
        return Para(StyleParaKind.HEADING_10)

    @property
    def horizontal_line(self) -> Para:
        """Style Horizontal Line"""
        return Para(StyleParaKind.HORIZONTAL_LINE)

    @property
    def idx(self) -> Para:
        """Style Index"""
        return Para(StyleParaKind.INDEX)

    @property
    def idx_bib1(self) -> Para:
        """Style Bibliography 1"""
        return Para(StyleParaKind.BIBLIOGRAPHY_1)

    @property
    def idx_c1(self) -> Para:
        """Style Contents 1"""
        return Para(StyleParaKind.CONTENTS_1)

    @property
    def idx_c2(self) -> Para:
        """Style Contents 2"""
        return Para(StyleParaKind.CONTENTS_2)

    @property
    def idx_c3(self) -> Para:
        """Style Contents 3"""
        return Para(StyleParaKind.CONTENTS_3)

    @property
    def idx_c4(self) -> Para:
        """Style Contents 4"""
        return Para(StyleParaKind.CONTENTS_4)

    @property
    def idx_c5(self) -> Para:
        """Style Contents 5"""
        return Para(StyleParaKind.CONTENTS_5)

    @property
    def idx_c6(self) -> Para:
        """Style Contents 6"""
        return Para(StyleParaKind.CONTENTS_6)

    @property
    def idx_c7(self) -> Para:
        """Style Contents 7"""
        return Para(StyleParaKind.CONTENTS_7)

    @property
    def idx_c8(self) -> Para:
        """Style Contents 8"""
        return Para(StyleParaKind.CONTENTS_8)

    @property
    def idx_c9(self) -> Para:
        """Style Contents 9"""
        return Para(StyleParaKind.CONTENTS_9)

    @property
    def idx_c10(self) -> Para:
        """Style Contents 10"""
        return Para(StyleParaKind.CONTENTS_10)

    @property
    def idx_separator(self) -> Para:
        """Style Index Separator"""
        return Para(StyleParaKind.INDEX_SEPARATOR)

    @property
    def idx_1(self) -> Para:
        """Style Index 1"""
        return Para(StyleParaKind.INDEX_1)

    @property
    def idx_2(self) -> Para:
        """Style Index 2"""
        return Para(StyleParaKind.INDEX_2)

    @property
    def idx_3(self) -> Para:
        """Style Index 3"""
        return Para(StyleParaKind.INDEX_3)

    @property
    def idx_obj1(self) -> Para:
        """Style Object Index 1"""
        return Para(StyleParaKind.OBJECT_INDEX_1)

    @property
    def idx_tbl1(self) -> Para:
        """Style Table Index 1"""
        return Para(StyleParaKind.TABLE_INDEX_1)

    @property
    def idx_user1(self) -> Para:
        """Style User Index 1"""
        return Para(StyleParaKind.USER_INDEX_1)

    @property
    def idx_user2(self) -> Para:
        """Style User Index 2"""
        return Para(StyleParaKind.USER_INDEX_2)

    @property
    def idx_user3(self) -> Para:
        """Style User Index 3"""
        return Para(StyleParaKind.USER_INDEX_3)

    @property
    def idx_user4(self) -> Para:
        """Style User Index 4"""
        return Para(StyleParaKind.USER_INDEX_4)

    @property
    def idx_user5(self) -> Para:
        """Style User Index 5"""
        return Para(StyleParaKind.USER_INDEX_5)

    @property
    def idx_user6(self) -> Para:
        """Style User Index 6"""
        return Para(StyleParaKind.USER_INDEX_6)

    @property
    def idx_user7(self) -> Para:
        """Style User Index 7"""
        return Para(StyleParaKind.USER_INDEX_7)

    @property
    def idx_user8(self) -> Para:
        """Style User Index 8"""
        return Para(StyleParaKind.USER_INDEX_8)

    @property
    def idx_user9(self) -> Para:
        """Style User Index 9"""
        return Para(StyleParaKind.USER_INDEX_9)

    @property
    def idx_user10(self) -> Para:
        """Style User Index 10"""
        return Para(StyleParaKind.USER_INDEX_10)

    @property
    def list_contents(self) -> Para:
        """Style List Contents"""
        return Para(StyleParaKind.LIST_CONTENTS)

    @property
    def list_heading(self) -> Para:
        """Style List Heading"""
        return Para(StyleParaKind.LIST_HEADING)

    @property
    def pre_text(self) -> Para:
        """Style Preformatted text"""
        return Para(StyleParaKind.PREFORMATTED_TEXT)

    @property
    def quotations(self) -> Para:
        """Style Quotations"""
        return Para(StyleParaKind.QUOTATIONS)

    @property
    def sender(self) -> Para:
        """Style Sender"""
        return Para(StyleParaKind.SENDER)

    @property
    def signature(self) -> Para:
        """Style Signature"""
        return Para(StyleParaKind.SIGNATURE)

    @property
    def tbl_contents(self) -> Para:
        """Style Table Contents"""
        return Para(StyleParaKind.TABLE_CONTENTS)

    @property
    def tbl_heading(self) -> Para:
        """Style Table Heading"""
        return Para(StyleParaKind.TABLE_HEADING)

    @property
    def txt_body(self) -> Para:
        """Style Text Body"""
        return Para(StyleParaKind.TEXT_BODY)

    @property
    def txt_first_line_indent(self) -> Para:
        """Style First Line Indent"""
        return Para(StyleParaKind.FIRST_LINE_INDENT)

    @property
    def txt_hanging_indent(self) -> Para:
        """Style Hanging Indent"""
        return Para(StyleParaKind.HANGING_INDENT)

    @property
    def txt_list_indent(self) -> Para:
        """Style List Indent"""
        return Para(StyleParaKind.LIST_INDENT)

    @property
    def txt_marginalia(self) -> Para:
        """Style Marginalia"""
        return Para(StyleParaKind.MARGINALIA)

    @property
    def txt_body_indent(self) -> Para:
        """Style Text Body Indent"""
        return Para(StyleParaKind.TEXT_BODY_INDENT)

    # endregion Style Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.STYLE | FormatKind.PARA
        return self._format_kind_prop

    @static_prop
    def default() -> Para:  # type: ignore[misc]
        """Gets ``Para`` default. Static Property."""
        try:
            return Para._DEFAULT_INST  # type: ignore[attr-defined]
        except AttributeError:
            Para._DEFAULT_INST = Para(name=StyleParaKind.STANDARD)  # type: ignore[attr-defined]
            Para._DEFAULT_INST._is_default_inst = True  # type: ignore[attr-defined]
        return Para._DEFAULT_INST  # type: ignore[attr-defined]
