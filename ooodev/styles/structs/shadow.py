"""
Module for Shadow format (``ShadowFormat``) struct.

.. versionadded:: 0.9.0
"""
# region imports
from __future__ import annotations
from typing import cast, overload

from ...meta.static_prop import static_prop
from ...utils import props as mProps
from ...utils.color import Color
from ...utils.color import CommonColor
from ..style_base import StyleBase

import uno
from ooo.dyn.table.shadow_format import ShadowFormat as ShadowFormat
from ooo.dyn.table.shadow_location import ShadowLocation as ShadowLocation


# endregion imports
class Shadow(StyleBase):
    """
    Shadow struct

    .. versionadded:: 0.9.0
    """

    # region init
    _EMPTY = None

    def __init__(
        self,
        location: ShadowLocation = ShadowLocation.BOTTOM_RIGHT,
        color: Color = CommonColor.GRAY,
        transparent: bool = False,
        width: float = 1.76,
    ) -> None:
        """
        Constructor

        Args:
            location (ShadowLocation, optional): contains the location of the shadow. Default to ``ShadowLocation.BOTTOM_RIGHT``.
            color (Color, optional):contains the color value of the shadow. Defaults to ``CommonColor.GRAY``.
            transparent (bool, optional): Shadow transparency. Defaults to False.
            width (float, optional): contains the size of the shadow (in mm units). Defaults to ``1.76``.

        Raises:
            ValueError: If ``color`` or ``width`` are less than zero.
        """
        if color < 0:
            raise ValueError("color must be a positive number")
        if width < 0:
            raise ValueError("Width must be a postivie number")
        init_vals = {
            "Location": location,
            "Color": color,
            "IsTransparent": transparent,
            "ShadowWidth": round(width * 100),
        }
        super().__init__(**init_vals)

    # endregion init

    # region methods

    def get_shadow_format(self) -> ShadowFormat:
        """
        Gets Shadow format for instance.

        Returns:
            ShadowFormat: Shadow Format
        """
        return ShadowFormat(
            Location=self._get("Location"),
            ShadowWidth=self._get("ShadowWidth"),
            IsTransparent=self._get("IsTransparent"),
            Color=self._get("Color"),
        )

    # region apply_style()

    @overload
    def apply_style(self, obj: object) -> None:
        ...

    def apply_style(self, obj: object, **kwargs) -> None:
        """
        Applies style to object

        Args:
            obj (object): Object that contains a ``ShadowFormat`` property.

        Returns:
            None:
        """
        shadow = self.get_shadow_format()
        mProps.Props.set(obj, ShadowFormat=shadow)

    # endregion apply_style()

    # endregion methods

    # region Properties

    @property
    def location(self) -> ShadowLocation:
        """Gets the location of the shadow."""
        return self._get("Location")

    @property
    def color(self) -> Color:
        """Gets the color value of the shadow."""
        return self._get("Color")

    @property
    def transparent(self) -> bool:
        """Gets transparent value"""
        return self._get("IsTransparent")

    @property
    def width(self) -> float:
        """Gets the size of the shadow (in mm units)"""
        pv = cast(int, self._get("ShadowWidth"))
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @static_prop
    def empty(cls) -> Shadow:
        """Gets empty Shadow. Static Property. when style is applied it remove any shadow."""
        if cls._EMPTY is None:
            cls._EMPTY = Shadow(location=ShadowLocation.NONE, transparent=False, color=8421504)
            # just to be exact due to float conversions.
            cls._EMPTY._set("ShadowWidth", 176)
        return cls._EMPTY

    # endregion Properties
