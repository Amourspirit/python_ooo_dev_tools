# region Imports
from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooo.dyn.drawing.hatch_style import HatchStyle

from ooodev.format.inner.direct.write.fill.area.hatch import Hatch as InnerHatch
from ooodev.format.inner.modify.write.page.page_style_base_multi import PageStyleBaseMulti
from ooodev.format.inner.preset.preset_hatch import PresetHatchKind
from ooodev.format.writer.style.page.kind.writer_style_page_kind import WriterStylePageKind
from ooodev.units.angle import Angle
from ooodev.utils.color import Color
from ooodev.utils.color import StandardColor

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

# endregion Imports


class Hatch(PageStyleBaseMulti):
    """
    Page Style Pattern

    .. seealso::

        - :ref:`help_writer_format_modify_page_area`

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
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
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
            bg_color(:py:data:`~.utils.color.Color`, optional): Specifies the background Color. Set this ``-1`` (default)
                for no background color.
            style_name (StyleParaKind, str, optional): Specifies the Page Style that instance applies to.
            Default is Default Page Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_page_area`
        """

        direct = InnerHatch(style=style, color=color, space=space, angle=angle, bg_color=bg_color)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> Hatch:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            Hatch: ``Hatch`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerHatch.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @classmethod
    def from_preset(
        cls,
        preset: PresetHatchKind,
        style_name: WriterStylePageKind | str = WriterStylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> Hatch:
        """
        Gets an instance from a preset.

        Args:
            preset (PresetHatchKind): Preset.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``PageStyles``.

        Returns:
            Hatch: ``Hatch`` instance from preset.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerHatch.from_preset(preset=preset)
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | WriterStylePageKind):
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
