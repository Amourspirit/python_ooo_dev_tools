from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.utils.helper.dot_dict import DotDict
from ooodev.loader import lo as mLo


if TYPE_CHECKING:
    from ooodev.calc.calc_doc import CalcDoc
    from ooodev.utils.data_type.range_obj import RangeObj


class PopupRngSelPartial:
    """A partial class for Selecting a range from a popup."""

    def __init__(self, doc: CalcDoc) -> None:
        self.__calc_doc = doc

    def invoke_range_selection(
        self,
        title: str = "Please select a range",
        close_on_mouse_release: bool = False,
        single_cell_mode: bool = False,
        initial_value: str = "",
    ) -> None:
        """
        Displays a range selection popup that allows the user to select a range with the mouse.

        If you are running from the command line, you can use the ``get_range_selection_from_popup()`` method instead.

        There is a automatic timeout of 60 seconds for the popup to be displayed.
        The timeout is to prevent the method from hanging indefinitely if the popup is not displayed.
        If the popup is not displayed within 60 seconds, the method will return ``None``.

        Args:
            title (str, optional): The title of the popup. Defaults to "Please select a range".
            close_on_mouse_release (bool, optional): Specifies if the dialog closes when mouse is released. Defaults to ``False``.
            single_cell_mode (bool, optional): Specifies if the dialog is in single cell mode. Defaults to ``False``.
            initial_value (str, optional): The initial value of the range. Defaults to "".

        Returns:
            None:

        Warning:
            This method requires the GUI to be present and will not work in Headless mode.

        Note:
            This method triggers the following events when this partial class is used in a class that inherits from EventsPartial:
                - BeforePopupRangeSelection
                - AfterPopupRangeSelection

            The event data for the BeforePopupRangeSelection event is a DotDict with the following keys:
                - doc: The ``CalcDoc`` object
                - title: The title of the popup
                - close_on_mouse_release: Specifies if the dialog closes when mouse is released.
                - single_cell_mode: Specifies if the dialog is in single cell mode.
                - initial_value: The initial value of the range.

            The event data for the AfterPopupRangeSelection event is a DotDict with the following keys:
                - view: The ``CalcSheetView`` object
                - state: The state of the selection, either "done" or "aborted"
                - rng_obj: The ``RangeObj`` object, which is the range selected by the user or ``None``.
                - close_on_mouse_release: Specifies if the dialog closes when mouse is released.
                - single_cell_mode: Specifies if the dialog is in single cell mode.
                - initial_value: The initial value of the range selection.
                - result: The result of the range selection from ``RangeSelectionEvent.RangeDescriptor`` Can be a string such as ``$Sheet1.$A$1:$B$2``.

            Because popup dialogs can block the main GUI Thread, this method is run in a separate thread.
            That means it is not possible to return the result of the range selection directly.
            Instead, the result is passed to the AfterPopupRangeSelection event and in a global event named ``GlobalCalcRangeSelector``.

            The ``GlobalCalcRangeSelector`` has the same event data as the ``AfterPopupRangeSelection`` event.

        Example:

            Example of using the global event ``GlobalCalcRangeSelector``.
            In this case ``MyObj`` could also be a dialog that need to be update when the range selection is done.

            .. code-block:: python

                from typing import Any
                from ooodev.globals import GblEvents
                from ooodev.events.args.event_args import EventArgs

                class MyObj:
                    def __init__(self):
                        self._fn_on_range_sel = self._on_range_sel
                        GblEvents().subscribe("GlobalCalcRangeSelector", self._fn_on_range_sel)

                    def on_range_selection(self, src:Any, event: EventArgs):
                        if event.event_data.state == "done":
                            print("Range Selection", event.event_data.rng_obj)

        .. versionadded:: 0.47.3
        """
        if mLo.Lo.bridge_connector.headless:
            self.__calc_doc.log.warning("Cannot invoke range selection in headless mode.")
            return None

        from ooodev.calc.sheet.range_selector import RangeSelectorThread

        cargs = None

        if isinstance(self, EventsPartial):
            cargs = CancelEventArgs(self)
            dd = DotDict(
                doc=self.__calc_doc,
                title=title,
                close_on_mouse_release=close_on_mouse_release,
                single_cell_mode=single_cell_mode,
                initial_value=initial_value,
            )
            cargs.event_data = dd
            self.trigger_event("BeforePopupRangeSelection", cargs)
            if cargs.cancel:
                return None

        # pylint: disable=import-outside-toplevel

        t1 = RangeSelectorThread(
            title=title,
            close_on_mouse_release=close_on_mouse_release,
            single_cell_mode=single_cell_mode,
            initial_value=initial_value,
        )
        if cargs:
            t1.add_event_observers(self.event_observer)  # type: ignore
        t1.start()
        return None

    def get_range_selection_from_popup(
        self,
        title: str = "Please select a range",
        close_on_mouse_release: bool = False,
        single_cell_mode: bool = False,
        initial_value: str = "",
    ) -> RangeObj | None:
        """
        Gets a range selection from a popup that allows the user to select a range with the mouse.

        There is a automatic timeout of 60 seconds for the popup to be displayed.
        The timeout is to prevent the method from hanging indefinitely if the popup is not displayed.
        If the popup is not displayed within 60 seconds, the method will return ``None``.

        If you are running from the command line, you can use this method; Otherwise, use ``invoke_range_selection()`` method instead.

        If macro mode ( no bridge connection ) is detected, the method will use the ``invoke_range_selection()`` method instead and no result will be returned.

        Args:
            title (str, optional): The title of the popup. Defaults to "Please select a range".
            close_on_mouse_release (bool, optional): Specifies if the dialog closes when mouse is released. Defaults to ``False``.
            single_cell_mode (bool, optional): Specifies if the dialog is in single cell mode. Defaults to ``False``.
            initial_value (str, optional): The initial value of the range. Defaults to "".

        Returns:
            RangeObj | None: The range object or ``None`` if no selection was made.

        Warning:
            This method requires the GUI to be present and will not work in Headless mode.

        Note:
            This method triggers the following events when this partial class is used in a class that inherits from EventsPartial:
                - BeforePopupRangeSelection
                - AfterPopupRangeSelection

            The event data for the BeforePopupRangeSelection event is a DotDict with the following keys:
                - doc: The ``CalcDoc`` object
                - title: The title of the popup
                - close_on_mouse_release: Specifies if the dialog closes when mouse is released.
                - single_cell_mode: Specifies if the dialog is in single cell mode.
                - initial_value: The initial value of the range.

            The event data for the AfterPopupRangeSelection event is a DotDict with the following keys:
                - view: The ``CalcSheetView`` object
                - state: The state of the selection, either "done" or "aborted"
                - rng_obj: The ``RangeObj`` object, which is the range selected by the user or ``None``.
                - close_on_mouse_release: Specifies if the dialog closes when mouse is released.
                - single_cell_mode: Specifies if the dialog is in single cell mode.
                - initial_value: The initial value of the range selection.
                - result: The result of the range selection from ``RangeSelectionEvent.RangeDescriptor`` Can be a string such as ``$Sheet1.$A$1:$B$2``.

            The ``GlobalCalcRangeSelector`` has the same event data as the ``AfterPopupRangeSelection`` event.
            See the ``invoke_range_selection()`` method for an example of using the global event.


        .. versionadded:: 0.47.1
        """
        if mLo.Lo.bridge_connector.headless:
            self.__calc_doc.log.warning("Cannot get range selection in headless mode.")
            return None

        if mLo.Lo.is_macro_mode:
            self.__calc_doc.log.warning(
                "Cannot get range selection in macro mode. Using invoke_range_selection instead."
            )
            self.invoke_range_selection(title, close_on_mouse_release, single_cell_mode, initial_value)
            return None
        from ooodev.calc.sheet.range_selector import RangeSelector

        cargs = None

        if isinstance(self, EventsPartial):
            cargs = CancelEventArgs(self)
            dd = DotDict(
                doc=self.__calc_doc,
                title=title,
                close_on_mouse_release=close_on_mouse_release,
                single_cell_mode=single_cell_mode,
                initial_value=initial_value,
            )
            cargs.event_data = dd
            self.trigger_event("BeforePopupRangeSelection", cargs)
            if cargs.cancel:
                return None

        # pylint: disable=import-outside-toplevel

        selector = RangeSelector(
            title=title,
            close_on_mouse_release=close_on_mouse_release,
            single_cell_mode=single_cell_mode,
            initial_value=initial_value,
        )
        if cargs:
            selector.add_event_observers(self.event_observer)  # type: ignore
        rng = selector.get_range_selection(doc=self.__calc_doc)

        return rng
