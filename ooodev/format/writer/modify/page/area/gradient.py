"""
Module for Page Style Fill Color Fill Color.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, overload

import uno
from ......exceptions import ex as mEx
from ......utils import info as mInfo
from ......utils import lo as mLo
from ......utils.color import Color
from ......utils.data_type.angle import Angle as Angle
from ......utils.data_type.offset import Offset as Offset
from ......utils.data_type.color_range import ColorRange as ColorRange
from ......utils.data_type.intensity_range import IntensityRange as IntensityRange
from ......utils.data_type.intensity import Intensity as Intensity
from .....direct.structs.gradient_struct import GradientStruct
from .....preset import preset_gradient
from .....preset.preset_gradient import PresetGradientKind as PresetGradientKind
from ....style.page.kind.style_page_kind import StylePageKind as StylePageKind
from ..page_style_base_multi import PageStyleBaseMulti

from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle
from ooo.dyn.drawing.fill_style import FillStyle

from com.sun.star.beans import XPropertySet


class FillStyleStruct(GradientStruct):
    def __init__(
        self,
        style: GradientStyle = GradientStyle.LINEAR,
        step_count: int = 0,
        x_offset: Intensity | int = 50,
        y_offset: Intensity | int = 50,
        angle: Angle | int = 0,
        border: Intensity | int = 0,
        start_color: Color = Color(0),
        start_intensity: Intensity | int = 100,
        end_color: Color = Color(16777215),
        end_intensity: Intensity | int = 100,
        style_name: StylePageKind | str = StylePageKind.STANDARD,
    ) -> None:
        """
        Constructor

        Args:
            style (GradientStyle, optional): Specifies the style of the gradient. Defaults to ``GradientStyle.LINEAR``.
            step_count (int, optional): Specifies the number of steps of change color. Defaults to ``0``.
            x_offset (Intensity, int, optional): Specifies the X-coordinate, where the gradient begins.
                This is effectively the center of the ``RADIAL``, ``ELLIPTICAL``, ``SQUARE`` and ``RECT`` style gradients. Defaults to ``50``.
            y_offset (Intensity, int, optional): Specifies the Y-coordinate, where the gradient begins.
                See: ``x_offset``. Defaults to ``50``.
            angle (Angle, int, optional): Specifies angle of the gradient. Defaults to 0.
            border (int, optional): Specifies percent of the total width where just the start color is used. Defaults to 0.
            start_color (Color, optional): Specifies the color at the start point of the gradient. Defaults to ``Color(0)``.
            start_intensity (Intensity, int, optional): Specifies the intensity at the start point of the gradient. Defaults to ``100``.
            end_color (Color, optional): Specifies the color at the end point of the gradient. Defaults to ``Color(16777215)``.
            end_intensity (Intensity, int, optional): Specifies the intensity at the end point of the gradient. Defaults to ``100``.
            style_name (StylePageKind, str, optional): Specifies the Page Style that instance applies to. Deftult is Default Page Style.
        """
        super().__init__(
            style=style,
            step_count=step_count,
            x_offset=x_offset,
            y_offset=y_offset,
            angle=angle,
            border=border,
            start_color=start_color,
            start_intensity=start_intensity,
            end_color=end_color,
            end_intensity=end_intensity,
        )
        self._style_name = str(style_name)

    def _get_property_name(self) -> str:
        return "FillGradient"

    def _supported_services(self) -> Tuple[str, ...]:
        return (
            "com.sun.star.style.PageStyle",
            "com.sun.star.style.Style",
        )

    def _is_valid_obj(self, obj: object) -> bool:
        valid = super()._is_valid_obj(obj)
        if valid:
            return True
        return mInfo.Info.is_doc_type(obj, mLo.Lo.Service.WRITER)

    # region apply()
    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies tab properties to ``obj``

        Args:
            obj (object): UNO Writer Document

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.format_named_event.FormatNamedEvent.STYLE_APPLYED` :eventref:`src-docs-event`

        Returns:
            None:
        """
        if not self._is_valid_obj(obj):
            # will not apply on this class but may apply on child classes
            self._print_not_valid_obj("apply")
            return
        p = self._get_style_props(obj)
        return super().apply(p)
        # GradinetStruct.apply(self, p)

    # endregion apply()
    def _get_style_props(self, obj: object) -> XPropertySet:
        return mInfo.Info.get_style_props(obj, "PageStyles", self.prop_style_name)

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StylePageKind):
        self._style_name = str(value)


class Gradient(PageStyleBaseMulti):
    """
    Page Fill Coloring

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        style: GradientStyle = GradientStyle.LINEAR,
        step_count: int = 0,
        offset: Offset = Offset(50, 50),
        angle: Angle | int = 0,
        border: Intensity | int = 0,
        grad_color: ColorRange = ColorRange(Color(0), Color(16777215)),
        grad_intensity: IntensityRange = IntensityRange(100, 100),
        name: str = "",
        style_name: StylePageKind | str = StylePageKind.STANDARD,
    ) -> None:
        """
        Constructor

        Args:
            style (GradientStyle, optional): Specifies the style of the gradient. Defaults to ``GradientStyle.LINEAR``.
            step_count (int, optional): Specifies the number of steps of change color. Defaults to ``0``.
            offset (Offset, int, optional): Specifies the X and Y coordinate, where the gradient begins.
                 X is is effectively the center of the ``RADIAL``, ``ELLIPTICAL``, ``SQUARE`` and ``RECT`` style gradients. Defaults to ``Offset(50, 50)``.
            angle (Angle, int, optional): Specifies angle of the gradient. Defaults to 0.
            border (int, optional): Specifies percent of the total width where just the start color is used. Defaults to 0.
            grad_color (ColorRange, optional): Specifies the color at the start point and stop point of the gradient. Defaults to ``ColorRange(Color(0), Color(16777215))``.
            grad_intensity (IntensityRange, optional): Specifies the intensity at the start point and stop point of the gradient. Defaults to ``IntensityRange(100, 100)``.
            name (str, optional): Specifies the Fill Gradient Name.
            style_name (StylePageKind, str, optional): Specifies the Page Style that instance applies to. Deftult is Default Page Style.

        Returns:
            None:
        """
        self._style_name = str(style_name)
        fs = FillStyleStruct(
            style=style,
            step_count=step_count,
            x_offset=offset.x,
            y_offset=offset.y,
            angle=angle,
            border=border,
            start_color=grad_color.start,
            start_intensity=grad_intensity.start,
            end_color=grad_color.end,
            end_intensity=grad_intensity.end,
            style_name=style_name,
        )
        super().__init__()

        self._set("FillStyle", FillStyle.GRADIENT)
        self._set("FillGradientStepCount", step_count)
        self._set("FillGradientName", self._get_gradient_name(style, name))
        self._set_style("fill_style", fs, *fs.get_attrs())

    def _get_gradient_name(self, style: GradientStyle, name: str) -> str:
        if name:
            return name
        if style == GradientStyle.AXIAL:
            return "Gradient 2"
        elif style == GradientStyle.ELLIPTICAL:
            return "Gradient 3"
        elif style == GradientStyle.SQUARE:
            # Square is quadratic
            return "Gradient 8"
        elif style == GradientStyle.RECT:
            # Rect is Square
            return "Gradient 7"
        else:
            return "gradient"

    @staticmethod
    def from_style(doc: object, style_name: StylePageKind | str = StylePageKind.STANDARD) -> Gradient:
        """
        Gets instance from object properties

        Args:
            obj (object): UNO Writer Document
            style_name (str, optional): Style to apply formating to. Default to the ``Default Page Style``.

        Raises:
            NotSupportedError: If ``obj`` is not a Writer Document.

        Returns:
            Gradient: Instance that represents Gradient Color.
        """
        bc = Gradient(style_name=style_name)
        if not bc._is_valid_obj(doc):
            raise mEx.NotSupportedError("obj is not a Writer Document")

        p = bc._get_style_props(doc)
        struct = GradientStruct.from_obj(p)
        prop_name = p.getPropertyValue("FillGradientName")
        inst = Gradient(
            style=struct.prop_style,
            step_count=struct.prop_step_count,
            offset=Offset(struct.prop_x_offset, struct.prop_y_offset),
            angle=struct.prop_angle,
            border=struct.prop_border,
            grad_color=ColorRange(struct.prop_start_color, struct.prop_end_color),
            grad_intensity=IntensityRange(int(struct.prop_start_intensity), int(struct.prop_end_intensity)),
            name=prop_name,
            style_name=style_name,
        )
        return inst

    @staticmethod
    def from_preset(preset: PresetGradientKind, style_name: StylePageKind | str = StylePageKind.STANDARD) -> Gradient:
        """
        Gets instance from preset

        Args:
            preset (PresetKind): Preset
            style_name (str, optional): Style to apply formating to. Default to the ``Default Page Style``.

        Returns:
            Gradient: Graident from a preset.
        """
        args = preset_gradient.get_preset(preset)
        args["style_name"] = style_name
        return Gradient(**args)

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StylePageKind):
        self._style_name = str(value)
