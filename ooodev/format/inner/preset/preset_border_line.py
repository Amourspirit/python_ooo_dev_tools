from __future__ import annotations
import typing

import uno
from ooo.dyn.drawing.line_joint import LineJoint
from ooo.dyn.drawing.line_style import LineStyle
from ooo.dyn.drawing.line_cap import LineCap
from ooo.dyn.drawing.line_dash import LineDash
from ooo.dyn.drawing.dash_style import DashStyle

from enum import Enum
from typing import NamedTuple


class BorderLineProps(NamedTuple):
    # use Union here for Sphinx autodoc and python < 3.10
    line_cap: LineCap
    line_dash_name: typing.Union[str, None]
    line_dash: typing.Union[str, None]
    line_joint: LineJoint
    line_style: typing.Union[LineStyle, None]


class BorderLineSeriesProps(NamedTuple):
    # use Union here for Sphinx autodoc and python < 3.10
    border_name: str
    border_style: LineStyle
    line_cap: LineCap
    line_dash: LineDash
    line_dash_name: typing.Union[str, None]
    line_style: LineStyle


# region Enum


class BorderLineKind(Enum):
    NONE = "None"
    CONTINUOUS = "Continuous"
    DOT = "Dot"
    DOT_ROUNDED = "Dot (Rounded)"
    LONG_DOT = "Long Dot"
    LONG_DOT_ROUNDED = "Long Dot (Rounded)"
    DASH = "Dash"
    DASH_ROUNDED = "Dash (Rounded)"
    LONG_DASH = "Long Dash"
    LONG_DASH_ROUNDED = "Long Dash (Rounded)"
    DOUBLE_DASH = "Double Dash"
    DOUBLE_DASH_ROUNDED = "Double Dash (Rounded)"
    DASH_DOT = "Dash Dot"
    DASH_DOT_ROUNDED = "Dash Dot (Rounded)"
    LONG_DASH_DOT = "Long Dash Dot"
    LONG_DASH_DOT_ROUNDED = "Long Dash Dot (Rounded)"
    DOUBLE_DASH_DOT = "Double Dash Dot"
    DOUBLE_DASH_DOT_ROUNDED = "Double Dash Dot (Rounded)"
    DASH_DOT_DOT = "Dash Dot Dot"
    DASH_DOT_DOT_ROUNDED = "Dash Dot Dot (Rounded)"
    DOUBLE_DASH_DOT_DOT = "Double Dash Dot Dot"
    DOUBLE_DASH_DOT_DOT_ROUNDED = "Double Dash Dot Dot (Rounded)"
    ULTRA_FINE_DOTTED = "Ultrafine Dotted (var)"
    FINE_DOTTED = "Fine Dotted"
    ULTRA_FINE_DASHED = "Ultrafine Dashed"
    FINE_DASHED = "Fine Dashed"
    DASHED = "Dashed (var)"
    LINE_STYLE_09 = "Line Style 9"
    DASHES_3_DOTS_3 = "3 Dashes 3 Dots (var)"
    ULTRA_FINE_DOTS_2_DASHES_3 = "Ultrafine 2 Dots 3 Dashes"
    DOTS_2_DASH_1 = "2 Dots 1 Dash"
    LINE_WITH_FINE_DOTS = "Line with Fine Dots"


# endregion Enum


# region border internal methods
def _get_line_kind_none() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name=None,
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=None,
    )


def _get_line_kind_continuous() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name=None,
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.SOLID,
    )


def _get_line_kind_dot() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name="Dot",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_dot_rounded() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.ROUND,
        line_dash_name="Dot (Rounded)",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_long_dot() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name="Long Dot",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_long_dot_rounded() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.ROUND,
        line_dash_name="Long Dot (Rounded)",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_dash() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name="Dash",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_dash_rounded() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.ROUND,
        line_dash_name="Dash (Rounded)",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_long_dash() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name="Long Dash",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_long_dash_rounded() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.ROUND,
        line_dash_name="Long Dash (Rounded)",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_double_dash() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name="Double Dash",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_double_dash_rounded() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.ROUND,
        line_dash_name="Double Dash (Rounded)",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_dash_dot() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name="Dash Dot",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_dash_dot_rounded() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.ROUND,
        line_dash_name="Dash Dot (Rounded)",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_long_dash_dot() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name="Long Dash Dot",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_long_dash_dot_rounded() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.ROUND,
        line_dash_name="Long Dash Dot (Rounded)",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_double_dash_dot() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name="Double Dash Dot",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_double_dash_dot_rounded() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.ROUND,
        line_dash_name="Double Dash Dot (Rounded)",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_dash_dot_dot() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name="Dash Dot Dot",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_dash_dot_dot_rounded() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.ROUND,
        line_dash_name="Dash Dot Dot (Rounded)",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_double_dash_dot_dot() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name="Double Dash Dot Dot",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_double_dash_dot_dot_rounded() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.ROUND,
        line_dash_name="Double Dash Dot Dot (Rounded)",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_ultra_fine_dotted() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name="Ultrafine Dotted (var)",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_fine_dotted() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name="Fine Dotted",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_ultra_fine_dashed() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name="Ultrafine Dashed",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_fine_dashed() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name="Fine Dashed",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_dashed() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name="Dashed (var)",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_line_style_09() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name="Line Style 9",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_dashes_3_dots_3() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name="3 Dashes 3 Dots (var)",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_ultra_fine_dots_2_dashes_3() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name="Ultrafine 2 Dots 3 Dashes",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_dots_2_dash_1() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name="2 Dots 1 Dash",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


def _get_line_kind_line_with_fine_dots() -> BorderLineProps:
    return BorderLineProps(
        line_cap=LineCap.BUTT,
        line_dash_name="Line with Fine Dots",
        line_dash=None,
        line_joint=LineJoint.ROUND,
        line_style=LineStyle.DASH,
    )


# endregion border internal methods


# region data series border internal methods
def _get_default_line_dash() -> LineDash:
    return LineDash(
        Style=DashStyle.RECT,  # type: ignore
        Dots=0,
        DotLen=0,
        Dashes=0,
        DashLen=0,
        Distance=0,
    )


def _get_series_line_kind_none() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="",
        border_style=LineStyle.NONE,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="",
        line_style=LineStyle.SOLID,
    )


def _get_series_line_kind_continuous() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Dot (Rounded)",
        border_style=LineStyle.SOLID,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Dot (Rounded)",
        line_style=LineStyle.SOLID,
    )


def _get_series_line_kind_dot() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Dot",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name=None,
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_dot_rounded() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Dot (Rounded)",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Dot",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_long_dot() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Long Dot",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Long Dot",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_long_dot_rounded() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Long Dot (Rounded)",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Long Dot (Rounded)",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_dash() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Dash",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Dash",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_dash_rounded() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Dash (Rounded)",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Dash (Rounded)",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_long_dash() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Long Dash",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Long Dash",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_long_dash_rounded() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Long Dash (Rounded)",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Long Dash (Rounded)",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_double_dash() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Double Dash",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Double Dash",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_double_dash_rounded() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Double Dash (Rounded)",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Double Dash (Rounded)",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_dash_dot() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Dash Dot",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Dash Dot",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_dash_dot_rounded() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Dash Dot (Rounded)",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Dash Dot (Rounded)",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_long_dash_dot() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Long Dash Dot",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Long Dash Dot",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_long_dash_dot_rounded() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Long Dash Dot (Rounded)",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Long Dash Dot (Rounded)",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_double_dash_dot() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Double Dash Dot",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Double Dash Dot",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_double_dash_dot_rounded() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Double Dash Dot (Rounded)",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Double Dash Dot (Rounded)",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_dash_dot_dot() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Dash Dot Dot",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Dash Dot Dot",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_dash_dot_dot_rounded() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Dash Dot Dot (Rounded)",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Dash Dot Dot (Rounded)",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_double_dash_dot_dot() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Double Dash Dot Dot",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Double Dash Dot Dot",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_double_dash_dot_dot_rounded() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Double Dash Dot Dot (Rounded)",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Double Dash Dot Dot (Rounded)",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_ultra_fine_dotted() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Ultrafine Dotted (var)",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Ultrafine Dotted (var)",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_fine_dotted() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Fine Dotted",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Fine Dotted",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_ultra_fine_dashed() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Ultrafine Dashed",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Ultrafine Dashed",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_fine_dashed() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Fine Dashed",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Fine Dashed",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_dashed() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Dashed (var)",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Dashed (var)",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_line_style_09() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Line Style 9",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Line Style 9",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_dashes_3_dots_3() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="3 Dashes 3 Dots (var)",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="3 Dashes 3 Dots (var)",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_ultra_fine_dots_2_dashes_3() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Ultrafine 2 Dots 3 Dashes",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Ultrafine 2 Dots 3 Dashes",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_dots_2_dash_1() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="2 Dots 1 Dash",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="2 Dots 1 Dash",
        line_style=LineStyle.DASH,
    )


def _get_series_line_kind_line_with_fine_dots() -> BorderLineSeriesProps:
    return BorderLineSeriesProps(
        border_name="Line with Fine Dots",
        border_style=LineStyle.DASH,
        line_cap=LineCap.BUTT,
        line_dash=_get_default_line_dash(),
        line_dash_name="Line with Fine Dots",
        line_style=LineStyle.DASH,
    )


# endregion data series border internal methods

# region Get Line Kind


def get_preset_border_line_props(kind: BorderLineKind) -> BorderLineProps:
    """
    Gets preset border line properties.

    Args:
        kind (PresetBorderLineKind): Preset border line kind.

    Returns:
        BorderLineProps: Preset border line properties.
    """
    # sourcery skip: low-code-quality
    if kind == BorderLineKind.CONTINUOUS:
        return _get_line_kind_continuous()
    elif kind == BorderLineKind.DOT:
        return _get_line_kind_dot()
    elif kind == BorderLineKind.DOT_ROUNDED:
        return _get_line_kind_dot_rounded()
    elif kind == BorderLineKind.LONG_DOT:
        return _get_line_kind_long_dot()
    elif kind == BorderLineKind.LONG_DOT_ROUNDED:
        return _get_line_kind_long_dot_rounded()
    elif kind == BorderLineKind.DASH:
        return _get_line_kind_dash()
    elif kind == BorderLineKind.DASH_ROUNDED:
        return _get_line_kind_dash_rounded()
    elif kind == BorderLineKind.LONG_DASH:
        return _get_line_kind_long_dash()
    elif kind == BorderLineKind.LONG_DASH_ROUNDED:
        return _get_line_kind_long_dash_rounded()
    elif kind == BorderLineKind.DOUBLE_DASH:
        return _get_line_kind_double_dash()
    elif kind == BorderLineKind.DOUBLE_DASH_ROUNDED:
        return _get_line_kind_double_dash_rounded()
    elif kind == BorderLineKind.DASH_DOT:
        return _get_line_kind_dash_dot()
    elif kind == BorderLineKind.DASH_DOT_ROUNDED:
        return _get_line_kind_dash_dot_rounded()
    elif kind == BorderLineKind.LONG_DASH_DOT:
        return _get_line_kind_long_dash_dot()
    elif kind == BorderLineKind.LONG_DASH_DOT_ROUNDED:
        return _get_line_kind_long_dash_dot_rounded()
    elif kind == BorderLineKind.DOUBLE_DASH_DOT:
        return _get_line_kind_double_dash_dot()
    elif kind == BorderLineKind.DOUBLE_DASH_DOT_ROUNDED:
        return _get_line_kind_double_dash_dot_rounded()
    elif kind == BorderLineKind.DASH_DOT_DOT:
        return _get_line_kind_dash_dot_dot()
    elif kind == BorderLineKind.DASH_DOT_DOT_ROUNDED:
        return _get_line_kind_dash_dot_dot_rounded()
    elif kind == BorderLineKind.DOUBLE_DASH_DOT_DOT:
        return _get_line_kind_double_dash_dot_dot()
    elif kind == BorderLineKind.DOUBLE_DASH_DOT_DOT_ROUNDED:
        return _get_line_kind_double_dash_dot_dot_rounded()
    elif kind == BorderLineKind.ULTRA_FINE_DOTTED:
        return _get_line_kind_ultra_fine_dotted()
    elif kind == BorderLineKind.FINE_DOTTED:
        return _get_line_kind_fine_dotted()
    elif kind == BorderLineKind.ULTRA_FINE_DASHED:
        return _get_line_kind_ultra_fine_dashed()
    elif kind == BorderLineKind.FINE_DASHED:
        return _get_line_kind_fine_dashed()
    elif kind == BorderLineKind.DASHED:
        return _get_line_kind_dashed()
    elif kind == BorderLineKind.LINE_STYLE_09:
        return _get_line_kind_line_style_09()
    elif kind == BorderLineKind.DASHES_3_DOTS_3:
        return _get_line_kind_dashes_3_dots_3()
    elif kind == BorderLineKind.ULTRA_FINE_DOTS_2_DASHES_3:
        return _get_line_kind_ultra_fine_dots_2_dashes_3()
    elif kind == BorderLineKind.DOTS_2_DASH_1:
        return _get_line_kind_dots_2_dash_1()
    elif kind == BorderLineKind.LINE_WITH_FINE_DOTS:
        return _get_line_kind_line_with_fine_dots()
    else:
        return _get_line_kind_none()


def get_preset_series_border_line_props(kind: BorderLineKind) -> BorderLineSeriesProps:
    """
    Gets preset series border line properties.

    This method is used to get the border line properties for the data series.

    Args:
        kind (PresetBorderLineKind): Preset border line kind.

    Returns:
        BorderLineProps: Preset border line properties.
    """
    # sourcery skip: low-code-quality
    if kind == BorderLineKind.CONTINUOUS:
        return _get_series_line_kind_continuous()
    elif kind == BorderLineKind.DOT:
        return _get_series_line_kind_dot()
    elif kind == BorderLineKind.DOT_ROUNDED:
        return _get_series_line_kind_dot_rounded()
    elif kind == BorderLineKind.LONG_DOT:
        return _get_series_line_kind_long_dot()
    elif kind == BorderLineKind.LONG_DOT_ROUNDED:
        return _get_series_line_kind_long_dot_rounded()
    elif kind == BorderLineKind.DASH:
        return _get_series_line_kind_dash()
    elif kind == BorderLineKind.DASH_ROUNDED:
        return _get_series_line_kind_dash_rounded()
    elif kind == BorderLineKind.LONG_DASH:
        return _get_series_line_kind_long_dash()
    elif kind == BorderLineKind.LONG_DASH_ROUNDED:
        return _get_series_line_kind_long_dash_rounded()
    elif kind == BorderLineKind.DOUBLE_DASH:
        return _get_series_line_kind_double_dash()
    elif kind == BorderLineKind.DOUBLE_DASH_ROUNDED:
        return _get_series_line_kind_double_dash_rounded()
    elif kind == BorderLineKind.DASH_DOT:
        return _get_series_line_kind_dash_dot()
    elif kind == BorderLineKind.DASH_DOT_ROUNDED:
        return _get_series_line_kind_dash_dot_rounded()
    elif kind == BorderLineKind.LONG_DASH_DOT:
        return _get_series_line_kind_long_dash_dot()
    elif kind == BorderLineKind.LONG_DASH_DOT_ROUNDED:
        return _get_series_line_kind_long_dash_dot_rounded()
    elif kind == BorderLineKind.DOUBLE_DASH_DOT:
        return _get_series_line_kind_double_dash_dot()
    elif kind == BorderLineKind.DOUBLE_DASH_DOT_ROUNDED:
        return _get_series_line_kind_double_dash_dot_rounded()
    elif kind == BorderLineKind.DASH_DOT_DOT:
        return _get_series_line_kind_dash_dot_dot()
    elif kind == BorderLineKind.DASH_DOT_DOT_ROUNDED:
        return _get_series_line_kind_dash_dot_dot_rounded()
    elif kind == BorderLineKind.DOUBLE_DASH_DOT_DOT:
        return _get_series_line_kind_double_dash_dot_dot()
    elif kind == BorderLineKind.DOUBLE_DASH_DOT_DOT_ROUNDED:
        return _get_series_line_kind_double_dash_dot_dot_rounded()
    elif kind == BorderLineKind.ULTRA_FINE_DOTTED:
        return _get_series_line_kind_ultra_fine_dotted()
    elif kind == BorderLineKind.FINE_DOTTED:
        return _get_series_line_kind_fine_dotted()
    elif kind == BorderLineKind.ULTRA_FINE_DASHED:
        return _get_series_line_kind_ultra_fine_dashed()
    elif kind == BorderLineKind.FINE_DASHED:
        return _get_series_line_kind_fine_dashed()
    elif kind == BorderLineKind.DASHED:
        return _get_series_line_kind_dashed()
    elif kind == BorderLineKind.LINE_STYLE_09:
        return _get_series_line_kind_line_style_09()
    elif kind == BorderLineKind.DASHES_3_DOTS_3:
        return _get_series_line_kind_dashes_3_dots_3()
    elif kind == BorderLineKind.ULTRA_FINE_DOTS_2_DASHES_3:
        return _get_series_line_kind_ultra_fine_dots_2_dashes_3()
    elif kind == BorderLineKind.DOTS_2_DASH_1:
        return _get_series_line_kind_dots_2_dash_1()
    elif kind == BorderLineKind.LINE_WITH_FINE_DOTS:
        return _get_series_line_kind_line_with_fine_dots()
    else:
        return _get_series_line_kind_none()


# endregion Get Line Kind
