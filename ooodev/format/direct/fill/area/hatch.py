"""
Module for Fill Properties Fill Hatch.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Any, Tuple, cast, overload

import uno

from .....events.args.key_val_cancel_args import KeyValCancelArgs
from .....events.format_named_event import FormatNamedEvent
from .....exceptions import ex as mEx
from .....utils import lo as mLo
from .....utils import props as mProps
from .....utils.color import Color
from .....utils.color import StandardColor
from .....utils.data_type.angle import Angle as Angle
from .....utils.unit_convert import UnitConvert
from ....kind.format_kind import FormatKind
from ....preset import preset_hatch as mPreset
from ....preset.preset_hatch import PresetHatchKind as PresetHatchKind
from ....style_base import StyleMulti
from ...structs.hatch_struct import HatchStruct
from .fill_color import FillColor


from ooo.dyn.drawing.fill_style import FillStyle
from ooo.dyn.drawing.hatch_style import HatchStyle as HatchStyle
from ooo.dyn.drawing.hatch import Hatch as UnoHatch


class Hatch(StyleMulti):
    """
    Class for Fill Properties Fill Hatch.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        style: HatchStyle = HatchStyle.SINGLE,
        color: Color = Color(0),
        space: float = 0.0,
        angle: Angle | int = 0,
        bg_color: Color = Color(-1),
    ) -> None:
        """
        Constructor

        Args:
            style (HatchStyle, optional): Specifies the kind of lines used to draw this hatch. Default ``HatchStyle.SINGLE``.
            color (Color, optional): Specifies the color of the hatch lines. Default ``0``.
            space (int, optional): Specifies the space between the lines in the hatch (in ``mm`` units). Default ``0.0``
            angle (Angle, int, optional): Specifies angle of the hatch in degrees. Default to ``0``.
            bg_color(Color, optionl): Specifies the background Color. Set this ``-1`` (default) for no background color.

        Returns:
            None:
        """

        hatch = HatchStruct(style=style, color=color, distance=space, angle=angle)
        bk_color = FillColor(bg_color)
        # FillStyle is set by this class
        bk_color._remove("FillStyle")
        # add event listener to prevent FillStyle from being set
        bk_color.add_event_listener(FormatNamedEvent.STYLE_SETTING, _on_bg_color_setting)
        bk_color.prop_color = bg_color

        init_vals = {}
        init_vals["FillStyle"] = FillStyle.HATCH
        if bg_color < 0:
            init_vals["FillBackground"] = False
        else:
            init_vals["FillBackground"] = True

        super().__init__(**init_vals)
        self._set_style("fill_color", bk_color, *bk_color.get_attrs())
        self._set_style("fill_hatch", hatch, *hatch.get_attrs())

    # region Internal Methods
    def _get_fill_color(self) -> FillColor:
        return self._get_style_inst("fill_color")

    def _get_hatch_struct(self) -> HatchStruct:
        return self._get_style_inst("fill_hatch")

    # endregion Internal Methods

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        return (
            "com.sun.star.drawing.FillProperties",
            "com.sun.star.text.TextContent",
            "com.sun.star.beans.PropertySet",
        )

    def _container_get_service_name(self) -> str:
        # https://github.com/LibreOffice/core/blob/d9e044f04ac11b76b9a3dac575f4e9155b67490e/chart2/source/tools/PropertyHelper.cxx#L229
        return "com.sun.star.drawing.HatchTable"

    # region apply()
    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO object that has supports ``com.sun.star.style.FillProperties`` or ``com.sun.star.beans.PropertySet`` service.

        Returns:
            None:
        """
        super().apply(obj, **kwargs)

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"Hatch.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()

    # endregion Overrides

    # region Static Methods
    @classmethod
    def from_preset(cls, preset: PresetHatchKind) -> Hatch:
        """
        Gets an instance from a preset

        Args:
            preset (PatternKind): Preset

        Returns:
            Hatch: Instance from preset.
        """
        kargs = mPreset.get_preset(preset)
        inst = super(Hatch, cls).__new__(cls)
        inst.__init__(**kargs)
        return inst

    @classmethod
    def from_obj(cls, obj: object) -> Hatch:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Hatch: ``Hatch`` instance that represents ``obj`` hatch pattern.
        """
        nu = super(Hatch, cls).__new__(cls)
        nu.__init__()
        if not nu._is_valid_obj(obj):
            raise mEx.NotSupportedError("Object is not support to convert to Hatch")

        hatch = cast(UnoHatch, mProps.Props.get(obj, "FillHatch"))
        fc = FillColor.from_obj(obj)

        inst = super(Hatch, cls).__new__(cls)
        inst.__init__(
            style=hatch.Style,
            color=hatch.Color,
            space=UnitConvert.convert_mm100_mm(hatch.Distance),
            angle=hatch.Angle,
            bg_color=fc.prop_color,
        )
        return inst

    # endregion Static Methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.TXT_CONTENT | FormatKind.FILL

    @property
    def prop_bg_color(self) -> Color:
        """Gets sets if fill image is tiled"""
        return self._get_fill_color().prop_color

    @prop_bg_color.setter
    def prop_bg_color(self, value: Color):
        fc = self._get_fill_color()
        fc.prop_color = value
        if value < 0:
            self._set("FillBackground", False)
        else:
            self._set("FillBackground", True)

    @property
    def prop_style(self) -> HatchStyle:
        """Gets/Sets the style of the hatch."""
        return self._get_hatch_struct().prop_style

    @prop_style.setter
    def prop_style(self, value: HatchStyle):
        hs = self._get_hatch_struct()
        hs.prop_style = value

    @property
    def prop_color(self) -> Color:
        """Gets/Sets the color of the hatch lines."""
        return self._get_hatch_struct().prop_color

    @prop_color.setter
    def prop_color(self, value: Color):
        hs = self._get_hatch_struct()
        hs.prop_color = value

    @property
    def prop_distance(self) -> float:
        """Gets/Sets the distance between the lines in the hatch (in ``mm`` units)."""
        return self._get_hatch_struct().prop_distance

    @prop_distance.setter
    def prop_distance(self, value: float):
        hs = self._get_hatch_struct()
        hs.prop_distance = value

    @property
    def prop_angle(self) -> Angle:
        """Gets/Sets angle of the hatch."""
        return self._get_hatch_struct().prop_angle

    @prop_angle.setter
    def prop_angle(self, value: Angle | int):
        hs = self._get_hatch_struct()
        hs.prop_angle = value

    # endregion Properties


def _on_bg_color_setting(source: Any, event_args: KeyValCancelArgs, *args, **kwargs) -> None:
    if event_args.key == "FillStyle":
        event_args.cancel = True
    elif event_args.key == "FillColor":
        # -1 means automatic color.
        # Fillcolor for hatch has not automatic color
        if event_args.value == -1:
            # strickly speaking this is not needed but follows how Draw handles it.
            event_args.value = StandardColor.DEFAULT_BLUE
