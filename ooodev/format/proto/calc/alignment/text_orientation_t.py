from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from ooodev.format.inner.direct.calc.alignment.text_orientation import EdgeKind
    from ooodev.units.angle import Angle
else:
    Protocol = object
    EdgeKind = Any
    Angle = Any

# See: ooodev.format.inner.direct.calc.alignment.text_orientation.TextOrientation


class TextOrientationT(StyleT, Protocol):
    """Cell Text Rotation."""

    def __init__(
        self, vert_stack: bool | None = ..., rotation: int | Angle | None = ..., edge: EdgeKind | None = ...
    ) -> None:
        """
        Constructor

        Args:
            vert_stack (bool, optional): Specifies if vertical stack is to be used.
            rotation (int, Angle, optional): Specifies if the rotation.
            edge (EdgeKind, optional): Specifies the Reference Edge.

        Returns:
            None:

        Note:
            When ``vert_stack`` is ``True`` other parameters are not used.

        Hint:
            - ``EdgeKind`` can be imported from ``ooodev.format.inner.direct.calc.alignment.text_orientation``
        """
        ...

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: object) -> TextOrientationT: ...

    @overload
    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TextOrientationT: ...

    # endregion from_obj()

    # region Properties

    @property
    def prop_vert_stacked(self) -> bool | None:
        """
        Gets/Sets vertically stacked.
        """
        ...

    @prop_vert_stacked.setter
    def prop_vert_stacked(self, value: bool | None): ...

    @property
    def prop_rotation(self) -> Angle | None:
        """Gets/Sets Vertical flip option"""
        ...

    @prop_rotation.setter
    def prop_rotation(self, value: int | Angle | None) -> None: ...

    @property
    def prop_edge(self) -> EdgeKind | None:
        """
        Gets/Sets Edge Kind.
        """
        ...

    @prop_edge.setter
    def prop_edge(self, value: EdgeKind | None): ...

    # endregion Properties
