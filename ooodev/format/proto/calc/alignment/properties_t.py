from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING

from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from ooodev.format.inner.direct.calc.alignment.properties import TextDirectionKind

else:
    Protocol = object
    TextDirectionKind = Any

# See: ooodev.format.inner.direct.calc.alignment.text_align.TextAlign


class PropertiesT(StyleT, Protocol):
    """Cell Text Properties."""

    def __init__(
        self,
        *,
        wrap_auto: bool | None = ...,
        hyphen_active: bool | None = ...,
        shrink_to_fit: bool | None = ...,
        direction: TextDirectionKind | None = ...,
    ) -> None:
        """
        Constructor

        Args:
            wrap_auto (bool, optional): Specifies wrap text automatically.
            hyphen_active (bool, optional): Specifies hyphenation active.
            shrink_to_fit (bool, optional): Specifies if text will shrink to cell.
            direction (TextDirectionKind, optional): Specifies Text Direction.

        Returns:
            None:

        Hint:
            - ``TextDirectionKind`` can be imported from ``ooodev.format.inner.direct.calc.alignment.properties``
        """
        ...

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: object) -> PropertiesT: ...

    @overload
    @classmethod
    def from_obj(cls, obj: object, **kwargs) -> PropertiesT: ...

    # endregion from_obj()

    # region Properties

    @property
    def prop_wrap_auto(self) -> bool | None:
        """
        Gets/Sets If text is wrapped automatically.
        """
        ...

    @prop_wrap_auto.setter
    def prop_wrap_auto(self, value: bool | None): ...

    @property
    def prop_hyphen_active(self) -> bool | None:
        """
        Gets/Sets If text is hyphenation is active.
        """
        ...

    @prop_hyphen_active.setter
    def prop_hyphen_active(self, value: bool | None): ...

    @property
    def prop_shrink_to_fit(self) -> bool | None:
        """
        Gets/Sets If text shrinks to cell size.
        """
        ...

    @prop_shrink_to_fit.setter
    def prop_shrink_to_fit(self, value: bool | None): ...

    @property
    def prop_direction(self) -> TextDirectionKind | None:
        """
        Gets/Sets Text Direction Kind.
        """
        ...

    @prop_direction.setter
    def prop_direction(self, value: TextDirectionKind | None): ...

    # endregion Properties
