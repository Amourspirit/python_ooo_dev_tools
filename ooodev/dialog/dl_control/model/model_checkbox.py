from __future__ import annotations
from typing import Any, TYPE_CHECKING
import contextlib
from pathlib import Path
from ooodev.units.unit_app_font_height import UnitAppFontHeight
from ooodev.units.unit_app_font_width import UnitAppFontWidth
from ooodev.units.unit_app_font_x import UnitAppFontX
from ooodev.units.unit_app_font_y import UnitAppFontY
from ooodev.adapter.awt.uno_control_check_box_model_partial import UnoControlCheckBoxModelPartial
from ooodev.utils.partial.model_prop_partial import ModelPropPartial
from ooodev.utils.file_io import FileIO
from ooodev.utils.type_var import PathOrStr

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlModel
    from ooodev.units.unit_obj import UnitT

# Model Position and Size are in AppFont units. View Size and Position are in Pixel units.


class ModelCheckbox(ModelPropPartial, UnoControlCheckBoxModelPartial):

    def __init__(self, model: UnoControlModel) -> None:
        """
        Constructor

        Args:
            component (UnoControlModel): UNO Component that implements ``com.sun.star.awt.UnoControlModel`` service.
        """
        ModelPropPartial.__init__(self, obj=model)  # must precede UnoControlButtonModelPartial
        UnoControlCheckBoxModelPartial.__init__(self, self.model)

    @property
    def context_writing_mode(self) -> int:
        """Get or set the context_writing_mode property."""
        return self.model.ContextWritingMode

    @context_writing_mode.setter
    def context_writing_mode(self, value: int) -> None:
        self.model.ContextWritingMode = value

    @property
    def enable_visible(self) -> bool:
        """Get or set the enable_visible property."""
        return self.model.EnableVisible

    @enable_visible.setter
    def enable_visible(self, value: bool) -> None:
        self.model.EnableVisible = value

    @property
    def height(self) -> UnitAppFontHeight:
        """Get the height of the dialog."""
        return UnitAppFontHeight(self.model.Height)

    @height.setter
    def height(self, value: float | UnitT) -> None:
        val = UnitAppFontHeight.from_unit_val(value)
        self.model.Height = int(val)

    # endregion UnoControlButtonModelPartial overrides

    @property
    def picture(self) -> str:
        """
        Gets/Sets the picture for the control

        When setting the value it can be a string or a Path object.
        If a string is passed it can be a URL or a path to a file.
        Value such as ``file:///path/to/image.png`` and ``/path/to/image.png`` are valid.
        Relative paths are supported.

        Returns:
            str: The picture URL in the format of ``file:///path/to/image.png`` or empty string if no picture is set.
        """
        with contextlib.suppress(Exception):
            return self.model.ImageURL
        return ""

    @picture.setter
    def picture(self, value: PathOrStr) -> None:
        pth_str = str(value)
        if not pth_str:
            self.model.ImageURL = ""
            return
        if isinstance(value, str):
            if value.startswith("file://"):
                self.model.ImageURL = value
                return
            value = Path(value)
        if not FileIO.is_valid_path_or_str(value):
            raise ValueError(f"Invalid path or str: {value}")
        self.model.ImageURL = FileIO.fnm_to_url(value)

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
    def x(self, value: float | UnitT) -> None:
        val = UnitAppFontX.from_unit_val(value)
        self.model.PositionX = int(val)

    @property
    def y(self) -> UnitAppFontY:
        """Get or set the position_y property."""
        return UnitAppFontY(self.model.PositionY)

    @y.setter
    def y(self, value: float | UnitT) -> None:
        val = UnitAppFontY.from_unit_val(value)
        self.model.PositionY = int(val)

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
    def tag(self) -> str:
        """Get or set the tag property."""
        return self.model.Tag

    @tag.setter
    def tag(self, value: str) -> None:
        self.model.Tag = value

    @property
    def width(self) -> UnitAppFontWidth:
        """Get or set the width property."""
        return UnitAppFontWidth(self.model.Width)

    @width.setter
    def width(self, value: float | UnitT) -> None:
        val = UnitAppFontWidth.from_unit_val(value)
        self.model.Width = int(val)

    if TYPE_CHECKING:

        @property
        def model(self) -> Any:
            """UnoControlModel Component"""
            return self.model
