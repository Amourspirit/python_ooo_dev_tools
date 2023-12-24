"""
Partial class for managing the applying of styles.

.. versionadded:: 0.17.9
"""
from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING

if TYPE_CHECKING:
    from ooodev.proto.style_obj import StyleT


class StylePartial:
    """
    Style methods

    .. versionadded:: 0.17.9
    """

    def __init__(self, component: Any):
        self.__component = component

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

    def apply_styles(self, *styles: StyleT, **kwargs) -> None:
        """
        Applies style to component.

        Args:
            styles expandable list of styles object such as ``Font`` to apply to ``obj``.
            kwargs (Any, optional): Expandable list of key value pairs.

        Returns:
            None:
        """
        for style in styles:
            style.apply(self.__component, **kwargs)
