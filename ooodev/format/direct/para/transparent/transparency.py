from __future__ import annotations
from typing import Tuple
import uno

from ....style_base import StyleMulti
from .....utils.data_type.angle import Angle as Angle
from .....utils.data_type.offset import Offset as Offset
from .....utils.data_type.intensity import Intensity as Intensity
from .....utils import gen_util as gUtil
from .....utils import props as mProps
from ....kind.format_kind import FormatKind
from .....events.args.key_val_cancel_args import KeyValCancelArgs
from ...fill.transparent.transparency import Transparency as FillTransparency


# PARA_BACK_COLOR_FLAGS = 0x7F000000
PARA_BACK_COLOR_FLAGS = 0x7C000000


class Transparency(StyleMulti):
    """
    Paragraph Transparency

    .. versionadded:: 0.9.0
    """

    def __init__(self, value: Intensity | int = 0) -> None:
        """
        Constructor

        Args:
            value (Intensity, int, optional): Specifies the transparency value from ``0`` to ``100``.
        """
        ft = FillTransparency(value)
        super().__init__()
        self._set("ParaBackColor", gUtil.NULL_OBJ)
        self._set_style("fill_transparency", ft, *ft.get_attrs())

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        return (
            "com.sun.star.drawing.FillProperties",
            "com.sun.star.text.TextContent",
        )

    def on_property_setting(self, event_args: KeyValCancelArgs) -> None:
        if event_args.key == "ParaBackColor":
            # event_args.cancel = True
            if event_args.value is gUtil.NULL_OBJ:
                event_args.value = -2071866342  # StandardColor.Lime result
                # color = cast(int, mProps.Props.get(event_args.event_data, "FillColor", None))
                # # color may be -1 automatic
                # if color is None or color < 0:
                #     event_args.value = PARA_BACK_COLOR_FLAGS
                # else:
                #     event_args.value = PARA_BACK_COLOR_FLAGS | color
        return super().on_property_setting(event_args)

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.TXT_CONTENT | FormatKind.FILL | FormatKind.PARA

    # endregion Properties

    # endregion Overrides
