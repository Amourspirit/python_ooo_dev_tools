from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.utils.helper.dot_dict import DotDict

if TYPE_CHECKING:
    from ooodev.calc.calc_doc import CalcDoc
    from ooodev.utils.data_type.range_obj import RangeObj


class PopupRngSelPartial:
    """A partial class for Selecting a range from a popup."""

    def __init__(self, doc: CalcDoc) -> None:
        self.__calc_doc = doc

    def get_range_selection_from_popup(
        self, title: str = "Please select a range", close_on_mouse_release: bool = False
    ) -> RangeObj | None:
        """
        Gets a range selection from a popup that allows the user to select a range with the mouse.

        Args:
            title (str, optional): The title of the popup. Defaults to "Please select a range".
            close_on_mouse_release (bool, optional): Specifies if the dialog closes when mouse is released. Defaults to ``False``.

        Returns:
            RangeObj | None: The range object or ``None`` if no selection was made.

        Warning:
            This method requires the GUI to be present and will not work in Headless mode.

        Note:
            This method triggers the following events when this partial class is used in a class that inherits from EventsPartial:
                - BeforePopupRangeSelection
                - AfterPopupRangeSelection

            The event data for the BeforePopupRangeSelection event is a DotDict with the following keys:
                - doc: The CalcDoc object
                - title: The title of the popup
                - close_on_mouse_release: Specifies if the dialog closes when mouse is released

            The event data for the AfterPopupRangeSelection event is a DotDict with the following keys:
                - doc: The CalcDoc object
                - rng_obj: The RangeObj object, which is the range selected by the user or None.


        .. versionadded:: 0.47.1
        """
        from ooodev.calc.sheet.range_selector import RangeSelector

        cargs = None

        if isinstance(self, EventsPartial):
            cargs = CancelEventArgs(self)
            dd = DotDict(doc=self.__calc_doc, title=title, close_on_mouse_release=close_on_mouse_release)
            cargs.event_data = dd
            self.trigger_event("BeforePopupRangeSelection", cargs)
            if cargs.cancel:
                return None

        # pylint: disable=import-outside-toplevel

        selector = RangeSelector(title=title, close_on_mouse_release=close_on_mouse_release)
        rng = selector.get_range_selection(doc=self.__calc_doc)
        if cargs is not None:
            eargs = EventArgs(self)
            dd = DotDict(doc=self.__calc_doc, rng_obj=rng)
            eargs.event_data = dd
            self.trigger_event("AfterPopupRangeSelection", eargs)  # type: ignore

        return rng
