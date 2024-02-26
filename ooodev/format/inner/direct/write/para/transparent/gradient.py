# region Imports
from __future__ import annotations
from typing import Any, Tuple

from ooo.dyn.awt.gradient_style import GradientStyle

from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.format.inner.direct.write.fill.transparent.gradient import Gradient as InnerGradient
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleMulti
from ooodev.loader import lo as mLo
from ooodev.units.angle import Angle
from ooodev.utils import gen_util as gUtil
from ooodev.utils.data_type.intensity import Intensity
from ooodev.utils.data_type.intensity_range import IntensityRange
from ooodev.utils.data_type.offset import Offset

# endregion Imports

# PARA_BACK_COLOR_FLAGS = 0x7F000000
PARA_BACK_COLOR_FLAGS = 0x7C000000


class Gradient(StyleMulti):
    """
    Paragraph Gradient Color

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        style: GradientStyle = GradientStyle.LINEAR,
        offset: Offset = Offset(50, 50),
        angle: Angle | int = 0,
        border: Intensity | int = 0,
        grad_intensity: IntensityRange = IntensityRange(0, 0),
    ) -> None:
        """
        Constructor

        Args:
            style (GradientStyle, optional): Specifies the style of the gradient. Defaults to ``GradientStyle.LINEAR``.
            step_count (int, optional): Specifies the number of steps of change color. Defaults to ``0``.
            offset (offset, optional): Specifies the X-coordinate (start) and Y-coordinate (end), where the gradient begins.
                X is effectively the center of the ``RADIAL``, ``ELLIPTICAL``, ``SQUARE`` and ``RECT`` style gradients. Defaults to ``Offset(50, 50)``.
            angle (Angle, int, optional): Specifies angle of the gradient. Defaults to 0.
            border (int, optional): Specifies percent of the total width where just the start color is used. Defaults to 0.
            grad_intensity (IntensityRange, optional): Specifies the intensity at the start point and stop point of the gradient. Defaults to ``IntensityRange(0, 0)``.
        """

        mLo.Lo.print("Gradient Transparency Class for Paragraph is not currently supported.")

        fg = InnerGradient(style=style, offset=offset, angle=angle, border=border, grad_intensity=grad_intensity)
        super().__init__()
        self._set("ParaBackColor", gUtil.NULL_OBJ)
        self._set_style("fill_grad", fg, *fg.get_attrs())  # type: ignore

    # region Overrides

    def _container_get_service_name(self) -> str:
        return "com.sun.star.drawing.TransparencyGradientTable"

    def _supported_services(self) -> Tuple[str, ...]:
        return (
            "com.sun.star.drawing.FillProperties",
            "com.sun.star.text.TextContent",
            "com.sun.star.style.ParagraphStyle",
        )

    def on_property_setting(self, source: Any, event_args: KeyValCancelArgs) -> None:
        """
        Triggers for each property that is set

        Args:
            source (Any): Event Source.
            event_args (KeyValueCancelArgs): Event Args
        """
        # sourcery skip: merge-nested-ifs
        if event_args.key == "ParaBackColor":
            # event_args.cancel = True
            if event_args.value is gUtil.NULL_OBJ:
                event_args.value = 2088883226
                # color = cast(int, mProps.Props.get(event_args.event_data, "FillColor", None))
                # # color may be -1 automatic
                # if color is None or color < 0:
                #     event_args.value = PARA_BACK_COLOR_FLAGS
                # else:
                #     event_args.value = PARA_BACK_COLOR_FLAGS | color
        super().on_property_setting(source, event_args)

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.TXT_CONTENT | FormatKind.FILL | FormatKind.PARA

    # endregion Properties

    # endregion Overrides
