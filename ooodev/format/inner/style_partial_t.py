from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING


if TYPE_CHECKING:
    from ooodev.proto.style_obj import StyleT
    from typing_extensions import Protocol
else:
    Protocol = object
    StyleT = Any


class StylePartialT(Protocol):
    """Type for StylePartial Class."""

    @overload
    def apply_styles(self, *styles: StyleT) -> None:
        """
        Applies style to component.

        Args:
            styles expandable list of styles object such as ``Font`` to apply to ``obj``.

        Returns:
            None:
        """
        ...

    @overload
    def apply_styles(self, *styles: StyleT, **kwargs) -> None:
        """
        Applies style to component.

        Args:
            styles expandable list of styles object such as ``Font`` to apply to ``obj``.
            kwargs (Any, optional): Expandable list of key value pairs.

        Returns:
            None:
        """
        ...
