from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.format.inner.partial.default_factor_styler import DefaultFactoryStyler
from ooodev.format.inner.style_factory import font_only_factory
from ooodev.events.partial.events_partial import EventsPartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.font.font_lang_t import FontLangT
    from ooodev.format.proto.font.font_only_t import FontOnlyT
    from ooodev.units.unit_obj import UnitT
else:
    FontLangT = Any
    FontOnlyT = Any
    UnitT = Any


class FontOnlyPartial:
    """
    Partial class for FontEffects.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        self.__styler = DefaultFactoryStyler(
            factory_name=factory_name,
            component=component,
            before_event="before_style_font_only",
            after_event="after_style_font_only",
            lo_inst=lo_inst,
        )
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)

    def style_font(
        self,
        name: str | None = None,
        size: float | UnitT | None = None,
        font_style: str | None = None,
        lang: FontLangT | None = None,
    ) -> FontOnlyT | None:
        """
        Style Font.

        Args:
            name (str | None, optional): Font Name.
            size (float | UnitT | None, optional): Font Size in ``PT`` units or ``UnitT``.
            font_style (str | None, optional): Font Style such as ``Bold Italics``
            lang (FontLangT | None, optional): Font Language.

        Raises:
            CancelEventError: If the event ``before_style_font_only`` is cancelled and not handled.

        Returns:
            FontOnlyT | None: Font Only instance or ``None`` if cancelled.

        See Also:
            :py:class:`~ooodev.format.inner.direct.write.char.font.font_only.FontLang`
        """
        kwargs = {}
        if name is not None:
            kwargs["name"] = name
        if size is not None:
            kwargs["size"] = size
        if font_style is not None:
            kwargs["font_style"] = font_style
        if lang is not None:
            kwargs["lang"] = lang
        return self.__styler.style(factory=font_only_factory, **kwargs)

    def style_font_get(self) -> FontOnlyT | None:
        """
        Gets the font Style.

        Raises:
            CancelEventError: If the event ``before_style_font_only_get`` is cancelled and not handled.

        Returns:
            FontOnlyT | None: Font style or ``None`` if cancelled.
        """
        return self.__styler.style_get(factory=font_only_factory)
