"""Module for Draw Style Fill Coloring."""

# region Import
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind as StyleParaKind
from ooodev.format.inner.modify.write.fill.fill_style_base_multi import FillStyleBaseMulti
from ooodev.format.draw.style.kind import DrawStyleFamilyKind
from ooodev.format.draw.style.lookup import FamilyGraphics
from ooodev.format.inner.direct.write.fill.area.hatch import Hatch as InnerHatch
from ooodev.format.inner.preset.preset_hatch import PresetHatchKind
from ooodev.utils.color import Color, StandardColor
from ooo.dyn.drawing.hatch_style import HatchStyle
from ooodev.units.angle import Angle


if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

# endregion Import


class Hatch(FillStyleBaseMulti):
    """
    Draw Style Fill Hatch

    .. seealso::

        - :ref:`help_draw_format_modify_area_hatch`

    .. versionadded:: 0.17.9
    """

    def __init__(
        self,
        *,
        style: HatchStyle = HatchStyle.SINGLE,
        color: Color = StandardColor.BLACK,
        space: float | UnitT = 0.0,
        angle: Angle | int = 0,
        bg_color: Color = StandardColor.AUTO_COLOR,
        style_name: str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> None:
        """
        Constructor

        Args:
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table then this property is required.
            name (str, optional): Specifies the name of the pattern. This is also the name that is used to store bitmap in LibreOffice Bitmap Table.
            tile (bool, optional): Specified if bitmap is tiled. Defaults to ``True``.
            stretch (bool, optional): Specifies if bitmap is stretched. Defaults to ``False``.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.
            style_name (FamilyGraphics, str, optional): Specifies the Style that instance applies to.
                Default is Default ``standard`` Style.
            style_family (str, DrawStyleFamilyKind, optional): Family Style. Defaults to ``graphics``.

        Returns:
            None:

        See Also:
            - :ref:`help_draw_format_modify_area_hatch`
        """

        direct = InnerHatch(
            style=style,
            color=color,
            space=space,
            angle=angle,
            bg_color=bg_color,
        )
        super().__init__()
        self._style_family_name = str(style_family)
        self._style_name = str(style_name)
        self._set_style("direct", direct)

    @classmethod
    def from_preset(
        cls,
        preset: PresetHatchKind,
        style_name: str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> Hatch:
        """
        Gets an instance from a preset

        Args:
            preset (PresetHatchKind): Preset

        Returns:
            Hatch: Instance from preset.
        """
        direct = InnerHatch.from_preset(preset)
        inst = cls(style_name=style_name, style_family=style_family)
        inst._set_style("direct", direct)
        return inst

    @classmethod
    def from_style(
        cls,
        doc: Any,
        style_name: FamilyGraphics | str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> Hatch:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (FamilyGraphics, str, optional): Specifies the Style that instance applies to.
                Default is ``FamilyGraphics.DEFAULT_DRAWING_STYLE``.
            style_family (DrawStyleFamilyKind, str, optional): Style family. Default ``DrawStyleFamilyKind.GRAPHICS``.

        Returns:
            Hatch: ``Hatch`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerHatch.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct)
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | FamilyGraphics):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerHatch:
        """Gets Inner instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerHatch, self._get_style_inst("direct"))
        return self._direct_inner
