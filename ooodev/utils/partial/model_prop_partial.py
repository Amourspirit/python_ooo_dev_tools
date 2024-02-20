from __future__ import annotations

from typing import Any, TYPE_CHECKING

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlModel  # service
else:
    UnoControl = Any


class ModelPropPartial:
    """A partial class for ``com.sun.star.awt.UnoControlModel`` child classes."""

    def __init__(self, obj: UnoControlModel) -> None:
        self.__ctl_model = obj

    @property
    def model(self) -> UnoControlModel:
        """Uno Control Model"""
        return self.__ctl_model
