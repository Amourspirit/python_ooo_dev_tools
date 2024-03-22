from __future__ import annotations
from typing import Any, TYPE_CHECKING
from ooodev.adapter.awt.uno_control_model_comp import UnoControlModelComp
from ooodev.units.unit_app_font_height import UnitAppFontHeight
from ooodev.units.unit_app_font_width import UnitAppFontWidth
from ooodev.units.unit_app_font_x import UnitAppFontX
from ooodev.units.unit_app_font_y import UnitAppFontY

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlModel
    from ooodev.units.unit_obj import UnitT
    from ooodev.utils.color import Color

# Model Position and Size are in AppFont units. View Size and Position are in Pixel units.


class ModelDialog(UnoControlModelComp):

    def __init__(self, component: UnoControlModel) -> None:
        """
        Constructor

        Args:
            component (UnoControlModel): UNO Component that implements ``com.sun.star.awt.UnoControlModel`` service.
        """
        UnoControlModelComp.__init__(self, component=component)

    @property
    def closeable(self) -> bool:
        """Get or set the closable property."""
        return self.model.Closeable

    @closeable.setter
    def closeable(self, value: bool) -> None:
        self.model.Closeable = value

    @property
    def background_color(self) -> Color:
        """Get or set the background color property."""
        return self.model.BackgroundColor

    @background_color.setter
    def background_color(self, value: Color) -> None:
        self.model.BackgroundColor = value

    @property
    def decoration(self) -> bool:
        """Get or set the decoration property."""
        return self.model.Decoration

    @decoration.setter
    def decoration(self, value: bool) -> None:
        self.model.Decoration = value

    @property
    def desktop_as_parent(self) -> bool:
        """Get or set the desktop_as_parent property."""
        return self.model.DesktopAsParent

    @desktop_as_parent.setter
    def desktop_as_parent(self, value: bool) -> None:
        self.model.DesktopAsParent = value

    @property
    def dialog_source_url(self) -> str:
        """Get or set the dialog_source_url property."""
        return self.model.DialogSourceURL

    @dialog_source_url.setter
    def dialog_source_url(self, value: str) -> None:
        self.model.DialogSourceURL = value

    @property
    def enabled(self) -> bool:
        """Get or set the enabled property."""
        return self.model.Enabled

    @enabled.setter
    def enabled(self, value: bool) -> None:
        self.model.Enabled = value

    @property
    def h_scroll(self) -> bool:
        """Get or set the h_scroll property."""
        return self.model.HScroll

    @h_scroll.setter
    def h_scroll(self, value: bool) -> None:
        self.model.HScroll = value

    @property
    def height(self) -> UnitAppFontHeight:
        """Get the height of the dialog."""
        return UnitAppFontHeight(self.model.Height)

    @height.setter
    def height(self, value: int | UnitT) -> None:
        val = UnitAppFontHeight.from_unit_val(value)
        self.model.Height = int(val)

    @property
    def help_text(self) -> str:
        """Get or set the help_text property."""
        return self.model.HelpText

    @help_text.setter
    def help_text(self, value: str) -> None:
        self.model.HelpText = value

    @property
    def help_url(self) -> str:
        """Get or set the help_url property."""
        return self.model.HelpURL

    @help_url.setter
    def help_url(self, value: str) -> None:
        self.model.HelpURL = value

    @property
    def image_url(self) -> str:
        """Get or set the image_url property."""
        return self.model.ImageURL

    @image_url.setter
    def image_url(self, value: str) -> None:
        self.model.ImageURL = value

    @property
    def moveable(self) -> bool:
        """Get or set the movable property."""
        return self.model.Moveable

    @moveable.setter
    def moveable(self, value: bool) -> None:
        self.model.Moveable = value

    @property
    def name(self) -> str:
        """Get or set the name property."""
        return self.model.Name

    @name.setter
    def name(self, value: str) -> None:
        self.model.Name = value

    @property
    def x(self) -> UnitAppFontX:
        """Get or set the position_x property."""
        return UnitAppFontX(self.model.PositionX)

    @x.setter
    def x(self, value: int | UnitT) -> None:
        val = UnitAppFontX.from_unit_val(value)
        self.model.PositionX = int(val)

    @property
    def y(self) -> UnitAppFontY:
        """Get or set the position_y property."""
        return UnitAppFontY(self.model.PositionY)

    @y.setter
    def y(self, value: int | UnitT) -> None:
        val = UnitAppFontY.from_unit_val(value)
        self.model.PositionY = int(val)

    @property
    def scroll_height(self) -> UnitAppFontHeight:
        """Get or set the scroll_bar_height property."""
        return UnitAppFontHeight(self.model.ScrollHeight)

    @scroll_height.setter
    def scroll_height(self, value: int | UnitT) -> None:
        val = UnitAppFontHeight.from_unit_val(value)
        self.model.ScrollHeight = int(val)

    @property
    def scroll_top(self) -> UnitAppFontX:
        """Get or set the scroll_top property."""
        return UnitAppFontX(self.model.ScrollTop)

    @scroll_top.setter
    def scroll_top(self, value: int | UnitT) -> None:
        val = UnitAppFontX.from_unit_val(value)
        self.model.ScrollTop = int(val)

    @property
    def scroll_width(self) -> UnitAppFontWidth:
        """Get or set the scroll_bar_width property."""
        return UnitAppFontWidth(self.model.ScrollWidth)

    @scroll_width.setter
    def scroll_width(self, value: int) -> None:
        val = UnitAppFontWidth.from_unit_val(value)
        self.model.ScrollWidth = int(val)

    @property
    def sizeable(self) -> bool:
        """Get or set the sizeable property."""
        return self.model.Sizeable

    @sizeable.setter
    def sizeable(self, value: bool) -> None:
        self.model.Sizeable = value

    @property
    def step(self) -> int:
        """Get or set the step property."""
        return self.model.Step

    @step.setter
    def step(self, value: int) -> None:
        self.model.Step = value

    @property
    def tab_index(self) -> int:
        """Get or set the tab_index property."""
        return self.model.TabIndex

    @tab_index.setter
    def tab_index(self, value: int) -> None:
        self.model.TabIndex = value

    @property
    def tab_page_id(self) -> int:
        """Get or set the tab_page_id property."""
        return self.model.TabPageID

    @tab_page_id.setter
    def tab_page_id(self, value: int) -> None:
        self.model.TabPageID = value

    @property
    def tag(self) -> str:
        """Get or set the tag property."""
        return self.model.Tag

    @tag.setter
    def tag(self, value: str) -> None:
        self.model.Tag = value

    @property
    def text_color(self) -> Color | None:
        """Get or set the text_color property."""
        return self.model.TextColor

    @text_color.setter
    def text_color(self, value: Color | None) -> None:
        self.model.TextColor = value

    @property
    def text_line_color(self) -> Color | None:
        """Get or set the text_line_color property."""
        return self.model.TextLineColor

    @text_line_color.setter
    def text_line_color(self, value: Color | None) -> None:
        self.model.TextLineColor = value

    @property
    def title(self) -> str:
        """Get or set the title property."""
        return self.model.Title

    @title.setter
    def title(self, value: str) -> None:
        self.model.Title = value

    @property
    def tool_tip(self) -> str:
        """Get or set the tool_tip property."""
        return self.model.ToolTip

    @tool_tip.setter
    def tool_tip(self, value: str) -> None:
        self.model.ToolTip = value

    @property
    def v_scroll(self) -> bool:
        """Get or set the v_scroll property."""
        return self.model.VScroll

    @v_scroll.setter
    def v_scroll(self, value: bool) -> None:
        self.model.VScroll = value

    @property
    def width(self) -> UnitAppFontWidth:
        """Get or set the width property."""
        return UnitAppFontWidth(self.model.Width)

    @width.setter
    def width(self, value: int | UnitT) -> None:
        val = UnitAppFontWidth.from_unit_val(value)
        self.model.Width = int(val)

    if TYPE_CHECKING:

        @property
        def model(self) -> Any:
            """UnoControlModel Component"""
            return self.model
