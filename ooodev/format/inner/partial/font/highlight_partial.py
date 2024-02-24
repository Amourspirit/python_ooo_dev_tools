from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.events.partial.events_partial import EventsPartial
from ooodev.format.inner.partial.factory_styler import FactoryStyler
from ooodev.format.inner.style_factory import font_highlight_factory
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.font.highlight_t import HighlightT
    from ooodev.utils.color import Color


class HighlightPartial:
    """
    Partial class for Font Highlight.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__styler = FactoryStyler(factory_name=factory_name, component=component, lo_inst=lo_inst)
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)
        self.__styler.after_event_name = "after_style_font_highlight"
        self.__styler.before_event_name = "before_style_font_highlight"

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
