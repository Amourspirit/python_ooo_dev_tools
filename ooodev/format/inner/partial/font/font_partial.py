from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from ooodev.format.inner.partial.default_factor_styler import DefaultFactoryStyler
from ooodev.format.inner.style_factory import font_factory
from ooodev.events.partial.events_partial import EventsPartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.write.char.font.font_t import FontT
    from ooo.dyn.awt.char_set import CharSetEnum
    from ooo.dyn.awt.font_family import FontFamilyEnum
    from ooo.dyn.awt.font_slant import FontSlant
    from ooo.dyn.awt.font_strikeout import FontStrikeoutEnum
    from ooo.dyn.awt.font_weight import FontWeightEnum
    from ooo.dyn.table.shadow_format import ShadowFormat

    from ooodev.format.inner.direct.write.char.font.font_effects import FontLine
    from ooodev.format.inner.direct.write.char.font.font_position import CharSpacingKind
    from ooodev.units.angle import Angle
    from ooodev.units.unit_obj import UnitT
    from ooodev.utils.color import Color
else:
    FontT = Any
    LoInst = Any


class FontPartial:
    """
    Partial class for General Font.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        self.__styler = DefaultFactoryStyler(
            factory_name=factory_name,
            component=component,
            before_event="before_style_general_font",
            after_event="after_style_general_font",
            lo_inst=lo_inst,
        )
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)

    def style_font_general(
        self,
        b: bool | None = None,
        i: bool | None = None,
        u: bool | None = None,
        bg_color: Color | None = None,
        bg_transparent: bool | None = None,
        charset: CharSetEnum | None = None,
        color: Color | None = None,
        family: FontFamilyEnum | None = None,
        name: str | None = None,
        overline: FontLine | None = None,
        rotation: int | Angle | None = None,
        shadow_fmt: ShadowFormat | None = None,
        shadowed: bool | None = None,
        size: float | UnitT | None = None,
        slant: FontSlant | None = None,
        spacing: CharSpacingKind | float | UnitT | None = None,
        strike: FontStrikeoutEnum | None = None,
        subscript: bool | None = None,
        superscript: bool | None = None,
        underline: FontLine | None = None,
        weight: FontWeightEnum | None = None,
        word_mode: bool | None = None,
    ) -> FontT | None:
        """
        Style Font.

        Args:
            b (bool, optional): Shortcut to set ``weight`` to bold.
            i (bool, optional): Shortcut to set ``slant`` to italic.
            u (bool, optional): Shortcut ot set ``underline`` to underline.
            bg_color (:py:data:`~.utils.color.Color`, optional): The value of the text background color.
            bg_transparent (bool, optional): Determines if the text background color is set to transparent.
            charset (CharSetEnum, optional): The text encoding of the font.
            color (:py:data:`~.utils.color.Color`, optional): The value of the text color. Setting to ``-1`` will cause automatic color.
            family (FontFamilyEnum, optional): Font Family.
            name (str, optional): This property specifies the name of the font style.
                It may contain more than one name separated by comma.
            overline (FontLine, optional): Character overline values.
            rotation (int, Angle, optional): Specifies the rotation of a character in degrees.
                Depending on the implementation only certain values may be allowed.
            shadow_fmt: (ShadowFormat, optional): Determines the type, color, and width of the shadow.
            shadowed (bool, optional): Specifies if the characters are formatted and displayed with a shadow effect.
            size (float, UnitT, optional): This value contains the size of the characters in ``pt`` (point) units
                or :ref:`proto_unit_obj`.
            slant (FontSlant, optional): The value of the posture of the document such as ``FontSlant.ITALIC``.
            spacing (CharSpacingKind, float, UnitT, optional): Specifies character spacing in ``pt`` (point) units
                or :ref:`proto_unit_obj`.
            strike (FontStrikeoutEnum, optional): Determines the type of the strike out of the character.
            subscript (bool, optional): Subscript option.
            superscript (bool, optional): Superscript option.
            underline (FontLine, optional): Character underline values.
            weight (FontWeightEnum, optional): The value of the font weight.
            word_mode(bool, optional): If ``True``, the underline and strike-through properties are not applied to
                white spaces.

        Raises:
            CancelEventError: If the event ``before_style_general_font`` is cancelled and not handled.

        Returns:
            FontT | None: Font instance or ``None`` if cancelled.


        Hint:
            - ``FontFamilyEnum`` can be imported from ``ooo.dyn.awt.font_family``
            - ``CharSetEnum`` can be imported from ``ooo.dyn.awt.char_set``
            - ``ShadowFormat`` can be imported from ``ooo.dyn.table.shadow_format``
            - ``FontSlant`` can be imported from ``ooo.dyn.awt.font_slant``
            - ``FontStrikeoutEnum`` can be imported from ``ooo.dyn.awt.font_strikeout``
            - ``FontWeightEnum`` can be imported from ``ooo.dyn.awt.font_weight``
            - ``FontLine`` can be imported from ``ooodev.format.inner.direct.write.char.font.font_effects``
            - ``CharSpacingKind`` can be imported from ``ooodev.format.inner.direct.write.char.font.font_position``
        """
        factory = font_factory
        kwargs = {
            "b": b,
            "i": i,
            "u": u,
            "bg_color": bg_color,
            "bg_transparent": bg_transparent,
            "charset": charset,
            "color": color,
            "family": family,
            "name": name,
            "overline": overline,
            "rotation": rotation,
            "shadow_fmt": shadow_fmt,
            "shadowed": shadowed,
            "size": size,
            "slant": slant,
            "spacing": spacing,
            "strike": strike,
            "subscript": subscript,
            "superscript": superscript,
            "underline": underline,
            "weight": weight,
            "word_mode": word_mode,
        }
        return self.__styler.style(factory=factory, **kwargs)
