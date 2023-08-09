"""Draw Fill Transparency"""
from __future__ import annotations
import uno
from ooodev.format.inner.direct.write.fill.transparent.transparency import Transparency as FillTransparency


class Transparency(FillTransparency):
    """
    Fill Transparency
    """

    def __init__(self, value: int = 0) -> None:
        """
        Constructor

        Args:
            value (Intensity, int, optional): Specifies the transparency value from ``0`` to ``100``.

        Returns:
            None:
        """
        super().__init__(value=value)
