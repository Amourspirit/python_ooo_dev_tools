from __future__ import annotations
from typing import TYPE_CHECKING, Tuple
import contextlib
import uno


from ooodev.units import UnitMM100
from ooodev.utils import info as mInfo
from ooodev.adapter.container.name_container_comp import NameContainerComp
from ooodev.adapter.container.index_replace_comp import IndexReplaceComp
from ooo.dyn.text.paragraph_vert_align import ParagraphVertAlignEnum

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue
    from com.sun.star.container import XIndexReplace
    from com.sun.star.container import XNameContainer
    from com.sun.star.graphic import XGraphic
    from com.sun.star.style import DropCapFormat  # struct
    from com.sun.star.style import LineSpacing  # struct
    from com.sun.star.style import ParagraphProperties
    from com.sun.star.style import TabStop
    from com.sun.star.style.BreakType import BreakTypeProto  # type: ignore
    from com.sun.star.style.GraphicLocation import GraphicLocationProto  # type: ignore
    from com.sun.star.style.ParagraphAdjust import ParagraphAdjustProto  # type: ignore
    from com.sun.star.table import BorderLine
    from com.sun.star.table import ShadowFormat  # struct
    from com.sun.star.util import Color  # type def
    from ooodev.units.unit_obj import UnitT


class ParagraphPropertiesPartial:
    """
    Partial class for ParagraphProperties.

    See Also:
        `API ParagraphProperties <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties.html>`_
    """

    def __init__(self, component: ParagraphProperties) -> None:
        """
        Constructor

        Args:
            component (ParagraphProperties): UNO Component that implements ``com.sun.star.style.ParagraphProperties`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``ParagraphProperties``.
        """
        self.__component = component

    # region ParagraphProperties
    @property
    def para_interop_grab_bag(self) -> Tuple[PropertyValue, ...] | None:
        """
        Gets/Sets grab bag of paragraph properties, used as a string-any map for interim interop purposes.

        This property is intentionally not handled by the ODF filter.
        Any member that should be handled there should be first moved out from this grab bag to a separate property.

        **optional**:
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaInteropGrabBag
        return None

    @para_interop_grab_bag.setter
    def para_interop_grab_bag(self, value: Tuple[PropertyValue, ...]) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaInteropGrabBag = value

    @property
    def para_tab_stops(self) -> Tuple[TabStop, ...] | None:
        """
        Gets/Sets the positions and kinds of the tab stops within this paragraph.

        **optional**:
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaTabStops
        return None

    @para_tab_stops.setter
    def para_tab_stops(self, value: Tuple[TabStop, ...]) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaTabStops = value

    @property
    def border_distance(self) -> UnitMM100 | None:
        """
        Gets/Sets the distance from the border to the object.

        When setting the value, it can be either a float or an instance of ``UnitT``.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return UnitMM100(self.__component.BorderDistance)
        return None

    @border_distance.setter
    def border_distance(self, value: float | UnitT) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.BorderDistance = UnitMM100.from_unit_val(value).value

    @property
    def bottom_border(self) -> BorderLine | None:
        """
        Gets/Sets the bottom border of the object.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.BottomBorder
        return None

    @bottom_border.setter
    def bottom_border(self, value: BorderLine) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.BottomBorder = value

    @property
    def bottom_border_distance(self) -> UnitMM100 | None:
        """
        Gets/Sets the distance from the bottom border to the object.

        When setting the value, it can be either a float or an instance of ``UnitT``.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return UnitMM100(self.__component.BottomBorderDistance)
        return None

    @bottom_border_distance.setter
    def bottom_border_distance(self, value: float | UnitT) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.BottomBorderDistance = UnitMM100.from_unit_val(value).value

    @property
    def break_type(self) -> BreakTypeProto | None:
        """
        Gets/Sets the type of break that is applied at the beginning of the table.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.BreakType
        return None

    @break_type.setter
    def break_type(self, value: BreakTypeProto) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.BreakType = value

    @property
    def continuing_previous_sub_tree(self) -> bool | None:
        """
        Gets that a child node of a parent node that is not counted is continuing the numbering of parent's previous node's sub tree.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ContinueingPreviousSubTree
        return None

    @property
    def drop_cap_char_style_name(self) -> str | None:
        """
        Gets/Sets the character style name for drop caps.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.DropCapCharStyleName
        return None

    @drop_cap_char_style_name.setter
    def drop_cap_char_style_name(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.DropCapCharStyleName = value

    @property
    def drop_cap_format(self) -> DropCapFormat | None:
        """
        Gets/Sets whether the first characters of the paragraph are displayed in capital letters and how they are formatted.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.DropCapFormat

    @drop_cap_format.setter
    def drop_cap_format(self, value: DropCapFormat) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.DropCapFormat = value

    @property
    def drop_cap_whole_word(self) -> bool | None:
        """
        Gets/Sets if the property DropCapFormat is applied to the whole first word.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.DropCapWholeWord
        return None

    @drop_cap_whole_word.setter
    def drop_cap_whole_word(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.DropCapWholeWord = value

    @property
    def left_border(self) -> BorderLine | None:
        """
        Gets/Sets the left border of the object.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.LeftBorder

    @left_border.setter
    def left_border(self, value: BorderLine) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.LeftBorder = value

    @property
    def left_border_distance(self) -> UnitMM100 | None:
        """
        Gets/Sets the distance from the left border to the object.

        When setting the value, it can be either a float or an instance of ``UnitT``.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return UnitMM100(self.__component.LeftBorderDistance)
        return None

    @left_border_distance.setter
    def left_border_distance(self, value: float | UnitT) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.LeftBorderDistance = UnitMM100.from_unit_val(value).value

    @property
    def list_id(self) -> str | None:
        """
        Gets/Sets the id of the list to which the paragraph belongs.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ListId
        return None

    @list_id.setter
    def list_id(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ListId = value

    @property
    def list_label_string(self) -> str | None:
        """
        Gets reading the generated numbering list label.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ListLabelString

    @property
    def numbering_is_number(self) -> bool | None:
        """
        Gets/Sets.

        Returns ``False`` if the paragraph is part of a numbering, but has no numbering label.

        A paragraph is part of a numbering, if a style for a numbering is set - see ``numbering_style_name``.
        If the paragraph is not part of a numbering the property is void.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.NumberingIsNumber
        return None

    @numbering_is_number.setter
    def numbering_is_number(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.NumberingIsNumber = value

    @property
    def numbering_level(self) -> int | None:
        """
        Gets/Sets the numbering level of the paragraph.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.NumberingLevel
        return None

    @numbering_level.setter
    def numbering_level(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.NumberingLevel = value

    @property
    def numbering_rules(self) -> IndexReplaceComp | None:
        """
        Gets/Sets the numbering rules applied to this paragraph.

        When setting the value, it can be either an instance of ``XIndexReplace`` or ``IndexReplaceComp``.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return IndexReplaceComp(self.__component.NumberingRules)

    @numbering_rules.setter
    def numbering_rules(self, value: XIndexReplace | IndexReplaceComp) -> None:
        with contextlib.suppress(AttributeError):
            if mInfo.Info.is_instance(value, IndexReplaceComp):
                self.__component.NumberingRules = value.component
            else:
                self.__component.NumberingRules = value  # type: ignore

    @property
    def numbering_start_value(self) -> int | None:
        """
        Gets/Sets the start value for numbering if a new numbering starts at this paragraph.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.NumberingStartValue
        return None

    @numbering_start_value.setter
    def numbering_start_value(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.NumberingStartValue = value

    @property
    def numbering_style_name(self) -> str | None:
        """
        Gets/Sets the name of the style for the numbering.

        The name must be one of the names which are available via ``XStyleFamiliesSupplier``.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.NumberingStyleName
        return None

    @numbering_style_name.setter
    def numbering_style_name(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.NumberingStyleName = value

    @property
    def outline_level(self) -> int | None:
        """
        Gets/Sets the outline level to which the paragraph belongs

        Value ``0`` indicates that the paragraph belongs to the body text.

        Values ``[1..10]`` indicates that the paragraph belongs to the corresponding outline level.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.OutlineLevel
        return None

    @outline_level.setter
    def outline_level(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.OutlineLevel = value

    @property
    def page_desc_name(self) -> str | None:
        """
        If this property is set, it creates a page break before the paragraph it belongs to and assigns the value as the name of the new page style sheet to use.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.PageDescName
        return None

    @page_desc_name.setter
    def page_desc_name(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.PageDescName = value

    @property
    def page_number_offset(self) -> int | None:
        """
        Gets/Sets if a page break property is set at a paragraph, this property contains the new value for the page number.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.PageNumberOffset
        return None

    @page_number_offset.setter
    def page_number_offset(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.PageNumberOffset = value

    @property
    def page_style_name(self) -> str | None:
        """
        Gets the name of the current page style.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.PageStyleName
        return None

    @property
    def para_adjust(self) -> ParagraphAdjustProto:
        """
        Gets/Sets the adjustment of a paragraph.
        """
        return self.__component.ParaAdjust

    @para_adjust.setter
    def para_adjust(self, value: ParagraphAdjustProto) -> None:
        self.__component.ParaAdjust = value

    @property
    def para_back_color(self) -> Color | None:
        """
        Gets/Sets the paragraph background color.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaBackColor
        return None

    @para_back_color.setter
    def para_back_color(self, value: Color) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaBackColor = value

    @property
    def para_back_graphic(self) -> XGraphic | None:
        """
        Gets/Sets the graphic for the background of a paragraph.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaBackGraphic
        return None

    @para_back_graphic.setter
    def para_back_graphic(self, value: XGraphic) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaBackGraphic = value

    @property
    def para_back_graphic_filter(self) -> str | None:
        """
        Gets/Sets the name of the graphic filter for the background graphic of a paragraph.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaBackGraphicFilter
        return None

    @para_back_graphic_filter.setter
    def para_back_graphic_filter(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaBackGraphicFilter = value

    @property
    def para_back_graphic_location(self) -> GraphicLocationProto | None:
        """
        Gets/Sets the value for the position of a background graphic.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaBackGraphicLocation

    @para_back_graphic_location.setter
    def para_back_graphic_location(self, value: GraphicLocationProto) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaBackGraphicLocation = value

    @property
    def para_back_graphic_url(self) -> str | None:
        """
        Gets/Sets the value of a link for the background graphic of a paragraph.

        Note the new behavior since it this was deprecated:
        This property can only be set and only external URLs are supported (no more ``vnd.sun.star.GraphicObject`` scheme).
        When an URL is set, then it will load the graphic and set the ParaBackGraphic property.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaBackGraphicURL
        return None

    @para_back_graphic_url.setter
    def para_back_graphic_url(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaBackGraphicURL = value

    @property
    def para_back_transparent(self) -> bool | None:
        """
        Gets/Sets if the paragraph background color is set to transparent.

        This value is ``True`` if the paragraph background color is set to transparent.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaBackTransparent
        return None

    @para_back_transparent.setter
    def para_back_transparent(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaBackTransparent = value

    @property
    def para_bottom_margin(self) -> UnitMM100:
        """
        Gets/Sets the bottom margin of the paragraph in ``100th mm``.

        The distance between two paragraphs is specified by:

        The greater one is chosen.

        This property accepts ``int`` and ``UnitT`` types when setting.
        """
        return UnitMM100(self.__component.ParaBottomMargin)

    @para_bottom_margin.setter
    def para_bottom_margin(self, value: int | UnitT) -> None:
        self.__component.ParaBottomMargin = UnitMM100.from_unit_val(value).value

    @property
    def para_context_margin(self) -> bool | None:
        """
        Gets/Sets if contextual spacing is used.

        If ``True``, the top and bottom margins of the paragraph should not be applied when the previous and next paragraphs have the same style.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaContextMargin
        return None

    @para_context_margin.setter
    def para_context_margin(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaContextMargin = value

    @property
    def para_expand_single_word(self) -> bool | None:
        """
        Gets/Sets if single words are stretched.

        It is only valid if ``para_adjust`` and ``para_last_line_adjust`` are also valid.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaExpandSingleWord
        return None

    @para_expand_single_word.setter
    def para_expand_single_word(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaExpandSingleWord = value

    @property
    def para_first_line_indent(self) -> int | None:
        """
        Gets/Sets the indent for the first line.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaFirstLineIndent

    @para_first_line_indent.setter
    def para_first_line_indent(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaFirstLineIndent = value

    @property
    def para_hyphenation_max_hyphens(self) -> int | None:
        """
        Gets/Sets the maximum number of consecutive hyphens.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaHyphenationMaxHyphens
        return None

    @para_hyphenation_max_hyphens.setter
    def para_hyphenation_max_hyphens(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaHyphenationMaxHyphens = value

    @property
    def para_hyphenation_max_leading_chars(self) -> int | None:
        """
        Gets/Sets the minimum number of characters to remain before the hyphen character (when hyphenation is applied).

        Note:
            Confusingly it is named Max but specifies a minimum.


        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaHyphenationMaxLeadingChars
        return None

    @para_hyphenation_max_leading_chars.setter
    def para_hyphenation_max_leading_chars(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaHyphenationMaxLeadingChars = value

    @property
    def para_hyphenation_max_trailing_chars(self) -> int:
        """
        Gets/Sets the minimum number of characters to remain after the hyphen character (when hyphenation is applied).

        Note:
            Confusingly it is named Max but specifies a minimum.
        """
        return self.__component.ParaHyphenationMaxTrailingChars

    @para_hyphenation_max_trailing_chars.setter
    def para_hyphenation_max_trailing_chars(self, value: int) -> None:
        self.__component.ParaHyphenationMaxTrailingChars = value

    @property
    def para_hyphenation_min_word_length(self) -> int | None:
        """
        Gets/Sets the minimum word length in characters, when hyphenation is applied.

        **since**

            LibreOffice 7.4

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaHyphenationMinWordLength
        return None

    @para_hyphenation_min_word_length.setter
    def para_hyphenation_min_word_length(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaHyphenationMinWordLength = value

    @property
    def para_hyphenation_no_caps(self) -> bool | None:
        """
        Specifies whether words written in CAPS will be hyphenated.

        Setting to true will disable hyphenation of words written in CAPS for this paragraph.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaHyphenationNoCaps
        return None

    @para_hyphenation_no_caps.setter
    def para_hyphenation_no_caps(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaHyphenationNoCaps = value

    @property
    def para_hyphenation_no_last_word(self) -> bool | None:
        """
        Specifies whether last word of paragraph will be hyphenated.

        Setting to true will disable hyphenation of last word for this paragraph.

        **since**

            LibreOffice 7.4

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaHyphenationNoLastWord
        return None

    @para_hyphenation_no_last_word.setter
    def para_hyphenation_no_last_word(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaHyphenationNoLastWord = value

    @property
    def para_hyphenation_zone(self) -> int | None:
        """
        Gets/Sets the hyphenation zone, i.e. allowed extra white space in the line before applying hyphenation.

        **since**

            LibreOffice 7.4

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaHyphenationZone
        return None

    @para_hyphenation_zone.setter
    def para_hyphenation_zone(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaHyphenationZone = value

    @property
    def para_is_auto_first_line_indent(self) -> bool | None:
        """
        Gets/Sets if the first line should be indented automatically.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaIsAutoFirstLineIndent
        return None

    @para_is_auto_first_line_indent.setter
    def para_is_auto_first_line_indent(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaIsAutoFirstLineIndent = value

    @property
    def para_is_connect_border(self) -> bool | None:
        """
        Gets/Sets if borders set at a paragraph are merged with the next paragraph.

        Borders are only merged if they are identical.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaIsConnectBorder
        return None

    @para_is_connect_border.setter
    def para_is_connect_border(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaIsConnectBorder = value

    @property
    def para_is_hyphenation(self) -> bool:
        """
        Gets/Sets if automatic hyphenation is applied.
        """
        return self.__component.ParaIsHyphenation

    @para_is_hyphenation.setter
    def para_is_hyphenation(self, value: bool) -> None:
        self.__component.ParaIsHyphenation = value

    @property
    def para_is_numbering_restart(self) -> bool | None:
        """
        Gets/Sets if the numbering rules restart, counting at the current paragraph.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaIsNumberingRestart
        return None

    @para_is_numbering_restart.setter
    def para_is_numbering_restart(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaIsNumberingRestart = value

    @property
    def para_keep_together(self) -> bool | None:
        """
        Gets/Sets if page or column breaks between this and the following paragraph are prevented.

        Setting this property to ``True`` prevents page or column breaks between this and the following paragraph.

        This feature is useful for preventing title paragraphs to be the last line on a page or column.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaKeepTogether
        return None

    @para_keep_together.setter
    def para_keep_together(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaKeepTogether = value

    @property
    def para_last_line_adjust(self) -> int:
        """
        Gets/Sets the adjustment of the last line.

        It is only valid if ``para_adjust`` is set to ``block``.
        """
        return self.__component.ParaLastLineAdjust

    @para_last_line_adjust.setter
    def para_last_line_adjust(self, value: int) -> None:
        self.__component.ParaLastLineAdjust = value

    @property
    def para_left_margin(self) -> UnitMM100:
        """
        Gets/Sets the left margin of the paragraph in `1/100th mm`` units.

        This property accepts ``int`` and ``UnitT`` types when setting.
        """
        return UnitMM100(self.__component.ParaLeftMargin)

    @para_left_margin.setter
    def para_left_margin(self, value: int) -> None:
        self.__component.ParaLeftMargin = UnitMM100.from_unit_val(value).value

    @property
    def para_line_number_count(self) -> bool | None:
        """
        Gets/Sets if the paragraph is included in the line numbering.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaLineNumberCount
        return None

    @para_line_number_count.setter
    def para_line_number_count(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaLineNumberCount = value

    @property
    def para_line_number_start_value(self) -> int | None:
        """
        Gets/Sets the start value for the line numbering.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaLineNumberStartValue
        return None

    @para_line_number_start_value.setter
    def para_line_number_start_value(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaLineNumberStartValue = value

    @property
    def para_line_spacing(self) -> LineSpacing | None:
        """
        Gets/Sets the type of the line spacing of a paragraph.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaLineSpacing
        return None

    @para_line_spacing.setter
    def para_line_spacing(self, value: LineSpacing) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaLineSpacing = value

    @property
    def para_orphans(self) -> int | None:
        """
        Gets/Sets the minimum number of lines of the paragraph that have to be at bottom of a page if the paragraph is spread over more than one page.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaOrphans
        return None

    @para_orphans.setter
    def para_orphans(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaOrphans = value

    @property
    def para_register_mode_active(self) -> bool | None:
        """
        Gets/Sets if the register mode is applied to a paragraph.

        Note:
            Register mode is only used if the register mode property of the page style is switched on.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaRegisterModeActive
        return None

    @para_register_mode_active.setter
    def para_register_mode_active(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaRegisterModeActive = value

    @property
    def para_right_margin(self) -> UnitMM100:
        """
        Gets/Sets the right margin of the paragraph in ``100th mm`` units.

        This property accepts ``float`` and ``UnitT`` types when setting.
        """
        return UnitMM100(self.__component.ParaRightMargin)

    @para_right_margin.setter
    def para_right_margin(self, value: float | UnitT) -> None:
        self.__component.ParaRightMargin = UnitMM100.from_unit_val(value).value

    @property
    def para_shadow_format(self) -> ShadowFormat | None:
        """
        Gets/Sets the type, color, and size of the shadow.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaShadowFormat

    @para_shadow_format.setter
    def para_shadow_format(self, value: ShadowFormat) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaShadowFormat = value

    @property
    def para_split(self) -> bool | None:
        """
        Gets/Sets - Setting this property to ``False`` prevents the paragraph from getting split into two pages or columns.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaSplit
        return None

    @para_split.setter
    def para_split(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaSplit = value

    @property
    def para_style_name(self) -> str | None:
        """
        Gets/Sets the name of the current paragraph style.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaStyleName
        return None

    @para_style_name.setter
    def para_style_name(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaStyleName = value

    @property
    def para_top_margin(self) -> UnitMM100:
        """
        determines the top margin of the paragraph in ``100th mm`` units.

        The distance between two paragraphs is specified by:

        The greater one is chosen.

        This property accepts ``float`` and ``UnitT`` types when setting.
        """
        return UnitMM100(self.__component.ParaTopMargin)

    @para_top_margin.setter
    def para_top_margin(self, value: float | UnitT) -> None:
        self.__component.ParaTopMargin = UnitMM100.from_unit_val(value).value

    @property
    def para_user_defined_attributes(self) -> NameContainerComp | None:
        """
        Gets/Sets - this property stores xml attributes.

        They will be saved to and restored from automatic styles inside xml files.

        Can be set with ``XNameContainer`` or ``NameContainerComp``.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return NameContainerComp(self.__component.ParaUserDefinedAttributes)
        return None

    @para_user_defined_attributes.setter
    def para_user_defined_attributes(self, value: XNameContainer | NameContainerComp) -> None:
        with contextlib.suppress(AttributeError):
            if mInfo.Info.is_instance(value, NameContainerComp):
                self.__component.ParaUserDefinedAttributes = value.component
            else:
                self.__component.ParaUserDefinedAttributes = value  # type: ignore

    @property
    def para_vert_alignment(self) -> ParagraphVertAlignEnum | None:
        """
        Gets/Set the vertical alignment of a paragraph.

        When setting the value, it can be either an integer or an instance of ``ParagraphVertAlignEnum``.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return ParagraphVertAlignEnum(self.__component.ParaVertAlignment)
        return None

    @para_vert_alignment.setter
    def para_vert_alignment(self, value: int | ParagraphVertAlignEnum) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaVertAlignment = ParagraphVertAlignEnum(value).value

    @property
    def para_widows(self) -> int | None:
        """
        Gets/Sets the minimum number of lines of the paragraph that have to be at top of a page if the paragraph is spread over more than one page.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.ParaWidows
        return None

    @para_widows.setter
    def para_widows(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.ParaWidows = value

    @property
    def right_border(self) -> BorderLine | None:
        """
        Gets/Sets the right border of the object.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.RightBorder
        return None

    @right_border.setter
    def right_border(self, value: BorderLine) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.RightBorder = value

    @property
    def right_border_distance(self) -> UnitMM100 | None:
        """
        Gets/Sets the distance from the right border to the object.

        When setting the value, it can be either an float or an instance of ``UnitT``.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return UnitMM100(self.__component.RightBorderDistance)
        return None

    @right_border_distance.setter
    def right_border_distance(self, value: float | UnitT) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.RightBorderDistance = UnitMM100.from_unit_val(value).value

    @property
    def top_border(self) -> BorderLine | None:
        """
        Gets/Sets the top border of the object.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.__component.TopBorder
        return None

    @top_border.setter
    def top_border(self, value: BorderLine) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.TopBorder = value

    @property
    def top_border_distance(self) -> UnitMM100 | None:
        """
        Gets/Sets the distance from the top border to the object.

        When setting the value, it can be either an float or an instance of ``UnitT``.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return UnitMM100(self.__component.TopBorderDistance)
        return None

    @top_border_distance.setter
    def top_border_distance(self, value: float | UnitT) -> None:
        with contextlib.suppress(AttributeError):
            self.__component.TopBorderDistance = UnitMM100.from_unit_val(value).value

    # endregion ParagraphProperties
