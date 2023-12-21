from __future__ import annotations
from typing import Set, Tuple
import uno
from ooo.dyn.drawing.text_animation_kind import TextAnimationKind

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase


class NoEffect(StyleBase):
    """
    This class represents the text Animation of an object that supports ``com.sun.star.drawing.TextProperties``.
    """

    def __init__(self) -> None:
        """
        Constructor.
        """
        super().__init__()
        self._set_animation()

    def _set_animation(self) -> None:
        self._set("TextAnimationKind", TextAnimationKind.NONE)

    def _get_uno_props(self) -> Set[str]:
        return {"TextAnimationKind"}

    # region Overridden Methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.drawing.TextProperties",)
        return self._supported_services_values

    # endregion Overridden Methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.SHAPE
        return self._format_kind_prop

    # endregion Properties
