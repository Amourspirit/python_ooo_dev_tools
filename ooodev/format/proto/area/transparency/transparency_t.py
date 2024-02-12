from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING
import uno


if TYPE_CHECKING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
    from ...style_t import StyleT

    from ooodev.utils.data_type.intensity import Intensity
else:
    Protocol = object
    Intensity = None


class TransparencyT(StyleT, Protocol):
    """Fill Area Transparency Protocol"""

    def __init__(self, value: Intensity | int = ...) -> None: ...

    @property
    def prop_value(self) -> Intensity:
        """Gets/Sets Transparency value"""
        ...

    @prop_value.setter
    def prop_value(self, value: Intensity | int) -> None: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> TransparencyT: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> TransparencyT: ...

    # region Properties
    @property
    def default(self) -> TransparencyT:  # type: ignore[misc]
        """Gets Transparency Default."""
        ...

    # endregion Properties
