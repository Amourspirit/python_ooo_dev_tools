from __future__ import annotations
from typing import cast

from ..utils import info as mInfo
from ..utils import lo as mLo
from ..exceptions import ex as mEx
from .style_base import StyleBase


class Padding(StyleBase):
    """Paragraph Padding"""

    # this class also set Borders Padding in borders.Border class.

    def __init__(
        self,
        left: float | None = None,
        right: float | None = None,
        top: float | None = None,
        bottom: float | None = None,
        padding_all: float | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (float, optional): Paragraph left padding (in mm units).
            right (float, optional): Paragraph right padding (in mm units).
            top (float, optional): Paragraph top padding (in mm units).
            bottom (float, optional): Paragraph bottom padding (in mm units).
            padding_all (float, optional): Paragraph left, right, top, bottom padding (in mm units). If argument is present then ``left``, ``right``, ``top``, and ``bottom`` arguments are ignored.

        Raises:
            ValueError: If any argument value is less than zero.

        Returns:
            None:
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        init_vals = {}

        def validate(val: float | None) -> None:
            if val is not None:
                if val < 0.0:
                    raise ValueError("padding values must be positive values")

        def set_val(key, value) -> None:
            nonlocal init_vals
            if not value is None:
                init_vals[key] = round(value * 100)

        validate(left)
        validate(right)
        validate(top)
        validate(bottom)
        validate(padding_all)
        para_attrs = ("ParaLeftMargin", "ParaRightMargin", "ParaTopMargin", "ParaBottomMargin")
        if padding_all is None:
            for key, value in zip(para_attrs, (left, right, top, bottom)):
                set_val(key, value)
        else:
            for key in para_attrs:
                set_val(key, padding_all)

        super().__init__(**init_vals)

    def apply_style(self, obj: object) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (object): Object that supports ``com.sun.star.style.ParagraphProperties`` service

        Returns:
            None:
        """
        if mInfo.Info.support_service(obj, "com.sun.star.style.ParagraphProperties"):
            try:
                super().apply_style(obj)
            except mEx.MultiError as e:
                mLo.Lo.print(f"Padding.apply_style(): Unable to set Property")
                for err in e.errors:
                    mLo.Lo.print(f"  {err}")
        else:
            mLo.Lo.print('Padding.apply_style(): "com.sun.star.style.ParagraphProperties" not supported')
        return None

    @property
    def left(self) -> float | None:
        """Gets paragraph left padding (in mm units)."""
        pv = cast(int, self._get("ParaLeftMargin"))
        if pv is None:
            return None
        if pv == 0.0:
            return 0.0
        return float(pv / 100)

    @property
    def right(self) -> float | None:
        """Gets paragraph right padding (in mm units)."""
        pv = cast(int, self._get("ParaRightMargin"))
        if pv is None:
            return None
        if pv == 0.0:
            return 0.0
        return float(pv / 100)

    @property
    def top(self) -> float | None:
        """Gets paragraph top padding (in mm units)."""
        pv = cast(int, self._get("ParaTopMargin"))
        if pv is None:
            return None
        if pv == 0.0:
            return 0.0
        return float(pv / 100)

    @property
    def bottom(self) -> float | None:
        """Gets paragraph bottom padding (in mm units)."""
        pv = cast(int, self._get("ParaBottomMargin"))
        if pv is None:
            return None
        if pv == 0.0:
            return 0.0
        return float(pv / 100)
