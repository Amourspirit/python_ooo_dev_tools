from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from ooodev.format.inner.direct.calc.alignment.text_align import VertAlignKind
    from ooodev.format.inner.direct.calc.alignment.text_align import HoriAlignKind
    from ooodev.units.unit_obj import UnitT
    from ooodev.units.unit_pt import UnitPT
else:
    Protocol = object
    VertAlignKind = Any
    HoriAlignKind = Any
    UnitT = Any
    UnitPT = Any

# See: ooodev.format.inner.direct.calc.alignment.text_align.TextAlign


class TextAlignT(StyleT, Protocol):
    """Cell Text Alignment."""

    def __init__(
        self,
        *,
        hori_align: HoriAlignKind | None = ...,
        indent: float | UnitT | None = ...,
        vert_align: VertAlignKind | None = ...,
    ) -> None:
        """
        Constructor

        Args:
            hori_align (HoriAlignKind, optional): Specifies Horizontal Alignment.
            indent: (float, UnitT, optional): Specifies indent in ``pt`` (point) units or :ref:`proto_unit_obj`.
                Only used when ``hori_align`` is set to ``HoriAlignKind.LEFT``
            vert_align (VertAdjustKind, optional): Specifies Vertical Alignment.

        Returns:
            None:

        Hint:
            - ``HoriAlignKind`` can be imported from ``ooodev.format.inner.direct.calc.alignment.text_align``
            - ``VertAlignKind`` can be imported from ``ooodev.format.inner.direct.calc.alignment.text_align``
        """
        ...

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: object) -> TextAlignT: ...

    @overload
    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> TextAlignT: ...

    # endregion from_obj()

    # region Properties

    @property
    def prop_hori_align(self) -> HoriAlignKind | None:
        """Gets/Sets Horizontal align value."""
        ...

    @prop_hori_align.setter
    def prop_hori_align(self, value: HoriAlignKind | None) -> None: ...

    @property
    def prop_indent(self) -> UnitPT | None:
        """
        Gets/Sets indent.
        """
        ...

    @prop_indent.setter
    def prop_indent(self, value: float | UnitT | None): ...

    @property
    def prop_vert_align(self) -> VertAlignKind | None:
        """Gets/Sets vertical align value."""
        ...

    @prop_vert_align.setter
    def prop_vert_align(self, value: VertAlignKind | None) -> None: ...

    # endregion Properties
