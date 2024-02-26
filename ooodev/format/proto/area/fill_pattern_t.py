from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING
import uno


from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from com.sun.star.awt import XBitmap
    from ooodev.format.inner.preset.preset_pattern import PresetPatternKind

else:
    Protocol = object
    XBitmap = Any
    PresetPatternKind = Any



class FillPatternT(StyleT, Protocol):
    """Fill Pattern Protocol"""

    def __init__(
        self,
        *,
        bitmap: XBitmap | None = None,
        name: str = "",
        tile: bool = True,
        stretch: bool = False,
        auto_name: bool = False,
    ) -> None:

        ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> FillPatternT: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> FillPatternT: ...

    @overload
    @classmethod
    def from_preset(cls, preset: PresetPatternKind) -> FillPatternT: ...

    @overload
    @classmethod
    def from_preset(cls, preset: PresetPatternKind, **kwargs) -> FillPatternT: ...

    # region Properties
    @property
    def prop_tile(self) -> bool:
        """Gets sets if fill image is tiled"""
        ...

    @prop_tile.setter
    def prop_tile(self, value: bool):
        ...

    @property
    def prop_stretch(self) -> bool:
        """Gets sets if fill image is stretched"""
        ...

    @prop_stretch.setter
    def prop_stretch(self, value: bool):
        ...
    # endregion Properties
