# region Import
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooo.dyn.drawing.hatch_style import HatchStyle

from ooodev.format.inner.direct.write.para.area.hatch import Hatch as InnerHatch
from ooodev.format.inner.modify.write.para.para_style_base_multi import ParaStyleBaseMulti
from ooodev.format.inner.preset import preset_hatch
from ooodev.format.inner.preset.preset_hatch import PresetHatchKind
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind
from ooodev.units.angle import Angle
from ooodev.utils.color import Color
from ooodev.utils.color import StandardColor

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT
# endregion Import


class Hatch(ParaStyleBaseMulti):
    """
    Paragraph Style Gradient Color

    .. seealso::

        - :ref:`help_writer_format_modify_para_hatch`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        style: HatchStyle = HatchStyle.SINGLE,
        color: Color = StandardColor.BLACK,
        space: float | UnitT = 0.0,
        angle: Angle | int = 0,
        bg_color: Color = StandardColor.AUTO_COLOR,
        name: str = "",
        auto_name: bool = False,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
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
            name (str, optional): Specifies the Hatch Name.
            auto_name (bool, optional): Specifies if Hatch is give an auto name such as ``Hatch ``. Default ``False``.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_para_hatch`
        """

        direct = InnerHatch(
            style=style, color=color, space=space, angle=angle, bg_color=bg_color, name=name, auto_name=auto_name
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @staticmethod
    def from_preset(
        preset: PresetHatchKind,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> Hatch:
        """
        Gets instance from preset

        Args:
            preset (PresetHatchKind): Preset

        Returns:
            Hatch: Hatch from a preset.
        """
        args = preset_hatch.get_preset(preset)
        args["name"] = str(preset)
        args["style_name"] = style_name
        args["style_family"] = style_family
        return Hatch(**args)

    @classmethod
    def from_style(
        cls,
        doc: Any,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> Hatch:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            Alignment: ``Alignment`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerHatch.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleParaKind):
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
        if not isinstance(value, InnerHatch):
            raise TypeError(f'Expected type of InnerHatch, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
