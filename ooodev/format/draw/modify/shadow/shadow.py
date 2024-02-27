"""
Draw Style Shadow.

.. versionadded:: 0.17.12
"""

from __future__ import annotations
from typing import cast, Any, TYPE_CHECKING
import uno

from ooodev.format.draw.style.kind import DrawStyleFamilyKind
from ooodev.format.draw.style.lookup import FamilyGraphics
from ooodev.format.inner.modify.draw.fill_properties_style_base_multi import FillPropertiesStyleBaseMulti
from ooodev.format.draw.direct.shadow.shadow import Shadow as InnerShadow

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT
    from ooodev.utils.data_type.intensity import Intensity
    from ooodev.utils.color import Color
    from ooodev.format.inner.direct.write.shape.area.shadow import ShadowLocationKind


class Shadow(FillPropertiesStyleBaseMulti):
    """
    Shadow Style value.

    .. seealso::

        - :ref:`help_draw_format_modify_shadow_shadow`

    .. versionadded:: 0.17.12
    """

    def __init__(
        self,
        *,
        use_shadow: bool | None = None,
        location: ShadowLocationKind | None = None,
        color: Color | None = None,
        distance: float | UnitT | None = None,
        blur: int | UnitT | None = None,
        transparency: int | Intensity | None = None,
        style_name: str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> None:
        """
        Constructor

        Args:
            use_shadow (bool, optional): Specifies if shadow is used.
            location (ShadowLocationKind , optional): Specifies the shadow location.
            color (Color , optional): Specifies shadow color.
            distance (float, UnitT , optional): Specifies shadow distance in ``mm`` units or :ref:`proto_unit_obj`.
            blur (int, UnitT, optional): Specifies shadow blur in ``pt`` units or in ``mm`` units  or :ref:`proto_unit_obj`.
            transparency (int , optional): Specifies shadow transparency value from ``0`` to ``100``.
            style_name (FamilyGraphics, str, optional): Specifies the Style that instance applies to.
                Default is Default ``standard`` Style.
            style_family (str, DrawStyleFamilyKind, optional): Family Style. Defaults to ``graphics``.

        Returns:
            None:

        See Also:
            - :ref:`help_draw_format_modify_shadow_shadow`
        """

        direct = InnerShadow(
            use_shadow=use_shadow,
            location=location,
            color=color,
            distance=distance,
            blur=blur,
            transparency=transparency,
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = str(style_family)
        self._set_style("direct", direct)

    @classmethod
    def from_style(
        cls,
        doc: Any,
        style_name: str = FamilyGraphics.DEFAULT_DRAWING_STYLE,
        style_family: str | DrawStyleFamilyKind = DrawStyleFamilyKind.GRAPHICS,
    ) -> Shadow:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (FamilyGraphics, str, optional): Specifies the Style that instance applies to.
                Default is ``FamilyGraphics.DEFAULT_DRAWING_STYLE``.
            style_family (DrawStyleFamilyKind, str, optional): Style family. Default ``DrawStyleFamilyKind.GRAPHICS``.

        Returns:
            Shadow: ``Shadow`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerShadow.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerShadow:
        """Gets/Sets Inner Font instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerShadow, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerShadow) -> None:
        if not isinstance(value, InnerShadow):
            raise TypeError(f'Expected type of InnerShadow, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
