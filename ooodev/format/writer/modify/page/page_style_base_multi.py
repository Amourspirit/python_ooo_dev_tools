"""
Base Class for Page Style.

.. versionadded:: 0.9.0
"""

# region Import
from __future__ import annotations
from typing import Any, Tuple, overload, TypeVar

from com.sun.star.beans import XPropertySet

from ooodev.exceptions import ex as mEx
from ooodev.utils import info as mInfo
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleMulti
from ooodev.format.writer.style.page.kind.writer_style_page_kind import WriterStylePageKind
from ooodev.events.format_named_event import FormatNamedEvent
from ooodev.events.args.event_args import EventArgs
from ooodev.events.args.cancel_event_args import CancelEventArgs

# endregion Import

# LibreOffice seems to have an unresolved bug with Background color.
# https://bugs.documentfoundation.org/show_bug.cgi?id=99125
# see Also: https://forum.openoffice.org/en/forum/viewtopic.php?p=417389&sid=17b21c173e4a420b667b45a2949b9cc5#p417389
# The solution to these issues is to apply FillColor to Paragraph cursors TextParagraph.

_TPageStyleBaseMulti = TypeVar("_TPageStyleBaseMulti", bound="PageStyleBaseMulti")


class PageStyleBaseMulti(StyleMulti):
    """
    Page Style Base

    .. versionadded:: 0.9.0
    """

    def _supported_services(self) -> Tuple[str, ...]:
        return ()

    def _is_valid_obj(self, obj: object) -> bool:
        return mInfo.Info.is_doc_type(obj, mLo.Lo.Service.WRITER)

    def copy(self: _TPageStyleBaseMulti) -> _TPageStyleBaseMulti:
        """Gets a copy of instance as a new instance"""
        cp = super().copy()
        cp.prop_style_name = self.prop_style_name
        return cp

    # region apply()

    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, **kwargs) -> None: ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (Any): UNO Writer Document

        Returns:
            None:
        """
        # sourcery skip: extract-method
        cargs = CancelEventArgs(source=f"{self.apply.__qualname__}")
        cargs.event_data = self
        if cargs.cancel:
            return
        self._events.trigger(FormatNamedEvent.STYLE_APPLYING, cargs)
        if cargs.cancel:
            return
        try:
            if self._is_valid_obj(obj):
                p = self._get_style_props(obj)
                # Could call p.setPropertyValue() here instead of Props.set()
                # but by calling Props.set() events are triggered.
                mProps.Props.set(p, **self._get_properties())
                eargs = EventArgs.from_args(cargs)
                self._events.trigger(FormatNamedEvent.STYLE_APPLIED, eargs)
                styles = self._get_multi_styles()
                for _, info in styles.items():
                    style, kw = info
                    if kw:
                        style.apply(obj, **kw.kwargs)
                    else:
                        style.apply(obj)
            else:
                mLo.Lo.print(f"{self.__class__.__name__}.apply(): Not a Writer Document. Unable to set Style Property")
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Style Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()

    def _get_style_props(self, obj: object) -> XPropertySet:
        return mInfo.Info.get_style_props(obj, "PageStyles", self.prop_style_name)

    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PAGE | FormatKind.STYLE

    @property
    def prop_style_name(self) -> str:
        """
        Gets/Sets property Style Name.

        Raises:
            NotImplementedError:
        """
        raise NotImplementedError

    @prop_style_name.setter
    def prop_style_name(self, value: str | WriterStylePageKind):
        raise NotImplementedError
