from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING
import contextlib
import uno

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.style_named_event import StyleNameEvent
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.mock import mock_g
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial

if TYPE_CHECKING:
    from ooo.dyn.style.paragraph_adjust import ParagraphAdjust
    from ooo.dyn.text.paragraph_vert_align import ParagraphVertAlignEnum
    from ooodev.format.inner.direct.write.para.align.writing_mode import WritingMode
    from ooodev.format.inner.direct.write.para.align.alignment import LastLineKind
    from ooodev.format.inner.direct.write.para.align.alignment import Alignment


class WriteParaAlignmentPartial:
    """
    Partial class for Write Paragraph Alignment.
    """

    def __init__(self, component: Any) -> None:
        self.__component = component

    def style_alignment(
        self,
        *,
        align: ParagraphAdjust | None = None,
        align_vert: ParagraphVertAlignEnum | None = None,
        txt_direction: WritingMode | None = None,
        align_last: LastLineKind | None = None,
        expand_single_word: bool | None = None,
        snap_to_grid: bool | None = None,
    ) -> Alignment | None:
        """
        Style Write Paragraph Alignment.

        Args:
            align (ParagraphAdjust, optional): Determines horizontal alignment of a paragraph.
            align_vert (ParagraphVertAlignEnum, optional): Determines vertical alignment of a paragraph.
            text_direction (WritingMode, optional): Determines the text direction.
            align_last (LastLineKind, optional): Determines the adjustment of the last line.
            expand_single_word (bool, optional): Determines if single words are stretched.
                It is only valid if ``align`` and ``align_last`` are also valid.
            snap_to_grid (bool, optional): Determines snap to text grid (if active).


        Raises:
            CancelEventError: If the event ``before_paragraph_alignment`` is cancelled and not handled.

        Returns:
            Alignment | None: Alignment Style instance or ``None`` if cancelled.

        Hint:
            - ``ParagraphAdjust`` can be imported from ``ooo.dyn.style.paragraph_adjust``
            - ``ParagraphVertAlignEnum`` can be imported from ``ooo.dyn.text.paragraph_vert_align``
            - ``LastLineKind`` can be imported from ``ooodev.format.writer.direct.para.alignment``
            - ``WritingMode`` can be imported from ``ooodev.format.inner.direct.write.para.align``
            - ``WritingMode2Enum`` can be imported from ``oo.dyn.text.writing_mode2``.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.format.inner.direct.write.para.align.alignment import Alignment

        comp = self.__component
        cargs = None
        if isinstance(self, EventsPartial):
            cargs = CancelEventArgs(self.style_alignment.__qualname__)
            event_data: Dict[str, Any] = {
                "align": align,
                "align_vert": align_vert,
                "txt_direction": txt_direction,
                "align_last": align_last,
                "expand_single_word": expand_single_word,
                "snap_to_grid": snap_to_grid,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event(StyleNameEvent.STYLE_APPLYING, cargs)
            self.trigger_event("before_paragraph_alignment", cargs)
            if cargs.cancel is True:
                if cargs.handled is True:
                    return None
                cargs.set("initial_event", "before_paragraph_alignment")
                self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                if cargs.handled is False:
                    raise mEx.CancelEventError(cargs, "Style has been cancelled.")
                else:
                    return None
            align = cargs.event_data.get("align", align)
            align_vert = cargs.event_data.get("align_vert", align_vert)
            txt_direction = cargs.event_data.get("txt_direction", txt_direction)
            align_last = cargs.event_data.get("align_last", align_last)
            expand_single_word = cargs.event_data.get("expand_single_word", expand_single_word)
            snap_to_grid = cargs.event_data.get("snap_to_grid", snap_to_grid)

            comp = cargs.event_data.get("this_component", comp)
        fe = Alignment(
            align=align,
            align_vert=align_vert,
            txt_direction=txt_direction,
            align_last=align_last,
            expand_single_word=expand_single_word,
            snap_to_grid=snap_to_grid,
        )
        if isinstance(self, LoInstPropsPartial):
            lo_inst = self.lo_inst
        else:
            lo_inst = mLo.Lo.current_lo
        with LoContext(lo_inst):
            fe.apply(comp)
        fe.set_update_obj(comp)
        if cargs is not None:
            # pylint: disable=no-member
            event_args = EventArgs.from_args(cargs)
            event_args.event_data["styler_object"] = fe
            self.trigger_event("after_paragraph_alignment", event_args)  # type: ignore
            self.trigger_event(StyleNameEvent.STYLE_APPLIED, event_args)  # type: ignore
        return fe

    def style_alignment_get(self) -> Alignment | None:
        """
        Gets the Alignment Style.

        Returns:
            Alignment | None: Alignment Style instance or ``None`` if not available.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.format.inner.direct.write.para.align.alignment import Alignment

        with contextlib.suppress(Exception):
            return Alignment.from_obj(self.__component)
        return None


if mock_g.FULL_IMPORT:
    from ooodev.format.inner.direct.write.para.align.alignment import Alignment
