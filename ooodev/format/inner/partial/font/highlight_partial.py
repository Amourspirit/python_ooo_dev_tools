from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.format.inner.partial.default_factor_styler import DefaultFactoryStyler
from ooodev.format.inner.style_factory import font_highlight_factory
from ooodev.events.partial.events_partial import EventsPartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.font.highlight_t import HighlightT
    from ooodev.utils.color import Color


class HighlightPartial:
    """
    Partial class for Font Highlight.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        self.__styler = DefaultFactoryStyler(
            factory_name=factory_name,
            component=component,
            before_event="before_style_font_highlight",
            after_event="after_style_font_highlight",
            lo_inst=lo_inst,
        )
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)

    def style_highlight(self, color: Color) -> HighlightT | None:
        """
        Style Font Highlight.

        Args:
            color (~ooodev.utils.color.Color, optional): Highlight Color. A value of ``-1`` Set color to Transparent.

        Raises:
            CancelEventError: If the event ``before_style_font_highlight`` is cancelled and not handled.

        Returns:
            HighlightT | None: Font Only instance or ``None`` if cancelled.

        See Also:
            :py:class:`~ooodev.format.inner.direct.write.char.font.font_only.FontLang`
        """
        return self.__styler.style(factory=font_highlight_factory, color=color)

    def style_highlight_get(self) -> HighlightT | None:
        """
        Gets the font highlight style.

        Raises:
            CancelEventError: If the event ``before_style_font_highlight_get`` is cancelled and not handled.

        Returns:
            HighlightT | None: Font style or ``None`` if cancelled.
        """
        return self.__styler.style_get(factory=font_highlight_factory)
