from __future__ import annotations
from typing import TYPE_CHECKING

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

        .. versionadded:: 0.47.1
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.calc.sheet.range_selector import RangeSelector

        selector = RangeSelector(title=title, close_on_mouse_release=close_on_mouse_release)
        rng = selector.get_range_selection(doc=self.__calc_doc)
        return rng
