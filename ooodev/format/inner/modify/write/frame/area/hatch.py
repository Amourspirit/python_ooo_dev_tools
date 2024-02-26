# region Imports
from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooo.dyn.drawing.hatch_style import HatchStyle

from ooodev.format.inner.direct.write.fill.area.hatch import Hatch as InnerHatch
from ooodev.format.inner.modify.write.frame.frame_style_base_multi import FrameStyleBaseMulti
from ooodev.format.inner.preset.preset_hatch import PresetHatchKind
from ooodev.format.writer.style.frame.style_frame_kind import StyleFrameKind
from ooodev.units.angle import Angle
from ooodev.utils.color import Color

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

# endregion Imports


class Hatch(FrameStyleBaseMulti):
    """
    Frame Style Hatch.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        style: HatchStyle = HatchStyle.SINGLE,
        color: Color = Color(0),
        space: float | UnitT = 0.0,
        angle: Angle | int = 0,
        bg_color: Color = Color(-1),
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> None:
        """
        Constructor

        Args:
            style (HatchStyle, optional): Specifies the kind of lines used to draw this hatch.
                Default ``HatchStyle.SINGLE``.
            color (:py:data:`~.utils.color.Color`, optional): Specifies the color of the hatch lines. Default ``0``.
            space (float, UnitT, optional): Specifies the space between the lines in the hatch (in ``mm`` units)
                or :ref:`proto_unit_obj`. Default ``0.0``
            angle (Angle, int, optional): Specifies angle of the hatch in degrees. Default to ``0``.
            bg_color(:py:data:`~.utils.color.Color`, optional): Specifies the background Color.
                Set this ``-1`` (default) for no background color.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Default is
                Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            None:
        """
        # pylint: disable=unexpected-keyword-arg
        direct = InnerHatch(
            style=style, color=color, space=space, angle=angle, bg_color=bg_color, _cattribs=self._get_inner_cattribs()  # type: ignore
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())  # type: ignore

    def _get_inner_cattribs(self) -> dict:
        return {"_supported_services_values": self._supported_services(), "_format_kind_prop": self.prop_format_kind}

    # region Static methods
    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> Hatch:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            Hatch: ``Hatch`` instance from style properties.
        """
        # pylint: disable=protected-access
        # pylint: disable=unexpected-keyword-arg
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerHatch.from_obj(obj=inst.get_style_props(doc), _cattribs=inst._get_inner_cattribs())
        direct._prop_parent = inst
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @classmethod
    def from_preset(
        cls,
        preset: PresetHatchKind,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> Hatch:
        """
        Gets an instance from a preset.

        Args:
            preset (PresetHatchKind): Preset.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            Hatch: ``Hatch`` instance from style properties.
        """
        # pylint: disable=protected-access
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerHatch.from_preset(preset=preset, _cattribs=inst._get_inner_cattribs())
        direct._prop_parent = inst
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    # endregion Static methods

    # region Properties
    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleFrameKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerHatch:
        """Gets/Sets Inner Hatch instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerHatch, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerHatch) -> None:
        # pylint: disable=protected-access
        # pylint: disable=unexpected-keyword-arg
        if not isinstance(value, InnerHatch):
            raise TypeError(f'Expected type of InnerHatch, got "{type(value).__name__}"')
        inst = value.__class__(
            style=value.prop_inner_hatch.prop_style,
            color=value.prop_inner_hatch.prop_color,
            space=value.prop_inner_hatch.prop_distance,
            angle=value.prop_inner_hatch.prop_angle,
            bg_color=value.prop_bg_color,
            _cattribs=self._get_inner_cattribs(),  # type: ignore
        )
        self._del_attribs("_direct_inner")
        self._set_style("direct", inst, *inst.get_attrs())  # type: ignore

    # endregion Properties
