# region Imports
from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooo.dyn.table.shadow_location import ShadowLocation

from ooodev.format.writer.style.frame.style_frame_kind import StyleFrameKind
from ooodev.format.inner.direct.write.para.border.shadow import Shadow as InnerShadow
from ooodev.utils.color import Color
from ooodev.utils.color import StandardColor
from ooodev.format.inner.modify.write.frame.frame_style_base_multi import FrameStyleBaseMulti

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

# endregion Imports


class Shadow(FrameStyleBaseMulti):
    """
    Frame Style Border Shadow.

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(
        self,
        *,
        location: ShadowLocation = ShadowLocation.BOTTOM_RIGHT,
        color: Color = StandardColor.GRAY,
        transparent: bool = False,
        width: float | UnitT = 1.76,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> None:
        """
        Constructor

        Args:
            location (ShadowLocation, optional): contains the location of the shadow. Default to ``ShadowLocation.BOTTOM_RIGHT``.
            color (:py:data:`~.utils.color.Color`, optional):contains the color value of the shadow. Defaults to ``StandardColor.GRAY``.
            transparent (bool, optional): Shadow transparency. Defaults to False.
            width (float, UnitT, optional): contains the size of the shadow (in ``mm`` units) or :ref:`proto_unit_obj`. Defaults to ``1.76``.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            None:
        """
        # pylint: disable=protected-access
        # pylint: disable=unexpected-keyword-arg

        direct = InnerShadow(
            location=location, color=color, transparent=transparent, width=width, _cattribs=self._get_inner_cattribs()  # type: ignore
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())  # type: ignore

    # endregion Init

    # region Static Methods
    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> Shadow:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            Shadow: ``Shadow`` instance from style properties.
        """
        # pylint: disable=protected-access
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerShadow.from_obj(obj=inst.get_style_props(doc), _cattribs=inst._get_inner_cattribs())
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    # endregion Static Methods

    # region internal methods
    def _get_inner_cattribs(self) -> dict:
        return {
            "_format_kind_prop": self.prop_format_kind,
            "_property_name": "ShadowFormat",
            "_supported_services_values": self._supported_services(),
        }

    # endregion internal methods

    # region Properties
    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleFrameKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerShadow:
        """Gets/Sets Inner Padding instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerShadow, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerShadow) -> None:
        if not isinstance(value, InnerShadow):
            raise TypeError(f'Expected type of InnerTransparency, got "{type(value).__name__}"')
        direct = value.copy(_cattribs=self._get_inner_cattribs())
        self._del_attribs("_direct_inner")
        self._set_style("direct", direct, *direct.get_attrs())

    # endregion Properties
