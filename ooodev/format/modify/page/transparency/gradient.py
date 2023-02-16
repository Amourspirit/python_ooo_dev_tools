from __future__ import annotations
from typing import Tuple, cast, Any
import uno
from ooo.dyn.awt.gradient_style import GradientStyle as GradientStyle

from .....utils.data_type.angle import Angle as Angle
from .....utils.data_type.offset import Offset as Offset
from .....utils.data_type.intensity_range import IntensityRange as IntensityRange
from ....writer.style.page.kind.style_page_kind import StylePageKind as StylePageKind
from ..page_style_base_multi import PageStyleBaseMulti
from .....utils.data_type.intensity import Intensity as Intensity

from ....direct.fill.transparent.gradient import Gradient as DirectGradient


class Gradient(PageStyleBaseMulti):
    """
    Page Style Transparency.

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
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> None:
        """
        Constructor

        Args:
            value (Intensity, int, optional): Specifies the transparency value from ``0`` to ``100``.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            None:
        """

        direct = DirectGradient(style=style, offset=offset, angle=angle, border=border, grad_intensity=grad_intensity)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StylePageKind | str = StylePageKind.STANDARD,
        style_family: str = "PageStyles",
    ) -> Gradient:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.
            style_family (str, optional): Style family. Defatult ``PageStyles``.

        Returns:
            Gradient: ``Gradient`` instance from document properties.
        """
        inst = super(Gradient, cls).__new__(cls)
        inst.__init__(style_name=style_name, style_family=style_family)
        direct = DirectGradient.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StylePageKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> DirectGradient:
        """Gets Inner Transparency instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(DirectGradient, self._get_style_inst("direct"))
        return self._direct_inner
