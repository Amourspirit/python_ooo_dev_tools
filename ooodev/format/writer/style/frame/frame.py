# region Import
from __future__ import annotations
from typing import Any, Tuple

from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.meta.static_prop import static_prop
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleName
from ooodev.format.writer.style.frame.style_frame_kind import StyleFrameKind

# endregion Import


class Frame(StyleName):
    """
    Frame Style.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together Border Table properties.

    .. seealso::

        - :ref:`help_writer_format_style_frame`

    .. versionadded:: 0.9.0
    """

    def __init__(self, name: StyleFrameKind | str = "") -> None:
        """
        Constructor

        Args:
            name (StyleFrameKind, str, optional): Style Name. Defaults to "Frame".

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_style_frame`
        """
        if name == "":
            name = Frame.default.prop_name
        super().__init__(name=name)

    # region Overrides

    def _get_family_style_name(self) -> str:
        return "FrameStyles"

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.text.TextFrame",)
        return self._supported_services_values

    def _get_property_name(self) -> str:
        try:
            return self._style_property_name
        except AttributeError:
            self._style_property_name = "FrameStyleName"
        return self._style_property_name

    def on_property_setting(self, source: Any, event_args: KeyValCancelArgs):
        """
        Triggers for each property that is set

        Args:
            event_args (KeyValueCancelArgs): Event Args
        """
        # there is only one style property for this class.
        # if CharStyleName is set to "" then an error is raised.
        # Solution is set to "No Character Style" or "Standard" Which LibreOffice recognizes and set to ""
        # this event covers apply() and restore()
        if event_args.value == "":
            event_args.value = Frame.default.prop_name
        super().on_property_setting(source, event_args)

    # endregion Overrides

    # region Style Properties
    @property
    def frame(self) -> Frame:
        """Style Frame"""
        return Frame(StyleFrameKind.FRAME)

    @property
    def formula(self) -> Frame:
        """Style Formula"""
        return Frame(StyleFrameKind.FORMULA)

    @property
    def graphics(self) -> Frame:
        """Style graphics"""
        return Frame(StyleFrameKind.GRAPHICS)

    @property
    def labels(self) -> Frame:
        """Style Labels"""
        return Frame(StyleFrameKind.LABELS)

    @property
    def marginalia(self) -> Frame:
        """Style Marginalia"""
        return Frame(StyleFrameKind.MARGINALIA)

    @property
    def OLE(self) -> Frame:
        """Style OLE"""
        return Frame(StyleFrameKind.OLE)

    @property
    def watermark(self) -> Frame:
        """Style Watermark"""
        return Frame(StyleFrameKind.WATERMARK)

    # endregion Style Properties

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.STYLE | FormatKind.FRAME
        return self._format_kind_prop

    @static_prop
    def default() -> Frame:  # type: ignore[misc]
        """Gets Frame default style. Static Property."""
        try:
            return Frame._DEFAULT_FRAME  # type: ignore[attr-defined]
        except AttributeError:
            Frame._DEFAULT_FRAME = Frame(name=StyleFrameKind.FRAME)  # type: ignore[attr-defined]
            Frame._DEFAULT_FRAME._is_default_inst = True  # type: ignore[attr-defined]
        return Frame._DEFAULT_FRAME  # type: ignore[attr-defined]

    # endregion Properties
