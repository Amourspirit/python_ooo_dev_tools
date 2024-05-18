from __future__ import annotations
from typing import Any, TYPE_CHECKING
import contextlib
from pathlib import Path
from ooodev.adapter.awt.uno_control_image_control_model_partial import UnoControlImageControlModelPartial
from ooodev.utils.partial.model_prop_partial import ModelPropPartial
from ooodev.utils.file_io import FileIO
from ooodev.utils.type_var import PathOrStr
from ooodev.dialog.dl_control.model.model_dialog_element_partial import ModelDialogElementPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlModel


class ModelImage(ModelPropPartial, UnoControlImageControlModelPartial, ModelDialogElementPartial):

    def __new__(cls, *args, **kwargs):
        return super().__new__(cls, *args, **kwargs)

    def __init__(self, model: UnoControlModel) -> None:
        """
        Constructor

        Args:
            component (UnoControlModel): UNO Component that implements ``com.sun.star.awt.UnoControlModel`` service.
        """
        ModelPropPartial.__init__(self, obj=model)  # must precede UnoControlButtonModelPartial
        UnoControlImageControlModelPartial.__init__(self, self.model)
        ModelDialogElementPartial.__init__(self, self.model)

    def __repr__(self) -> str:
        if hasattr(self, "name"):
            return f"ModelImage({self.name})"
        return "ModelImage"

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

    if TYPE_CHECKING:

        @property
        def model(self) -> Any:
            """UnoControlModel Component"""
            return self.model
