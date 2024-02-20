from __future__ import annotations

from typing import Any, TYPE_CHECKING

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControl  # service
else:
    UnoControl = Any


class ViewPropPartial:
    """A partial class for ``com.sun.star.awt.UnoControl`` child classes."""

    def __init__(self, obj: UnoControl) -> None:
        self.__ctl_view = obj

    @property
    def view(self) -> UnoControl:
        """Uno Control"""
        return self.__ctl_view
