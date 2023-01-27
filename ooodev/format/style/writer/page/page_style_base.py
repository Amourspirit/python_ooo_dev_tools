"""
Module for Page Style Fill Color Fill Color.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, overload

import uno
from .....events.event_singleton import _Events
from .....utils import lo as mLo
from .....utils import props as mProps
from .....utils import info as mInfo
from .....utils.type_var import T
from .....exceptions import ex as mEx
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase, EventArgs, CancelEventArgs, FormatNamedEvent

from com.sun.star.beans import XPropertySet

# LibreOffice seems to have an unresolved bug with Background color.
# https://bugs.documentfoundation.org/show_bug.cgi?id=99125
# see Also: https://forum.openoffice.org/en/forum/viewtopic.php?p=417389&sid=17b21c173e4a420b667b45a2949b9cc5#p417389
# The solution to these issues is to apply FillColor to Paragraph cursors TextParagraph.


class PageStyleBase(StyleBase):
    """
    Page Style Base

    .. versionadded:: 0.9.0
    """

    def _supported_services(self) -> Tuple[str, ...]:
        return ()

    def _is_valid_obj(self, obj: object) -> bool:
        return mInfo.Info.is_doc_type(obj, mLo.Lo.Service.WRITER)

    def copy(self: T) -> T:
        """Gets a copy of instance as a new instance"""
        cp = super().copy()
        cp.prop_style_name = self.prop_style_name

    # region apply()

    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (object): UNO Writer Document

        Returns:
            None:
        """
        cargs = CancelEventArgs(source=f"{self.apply.__qualname__}")
        cargs.event_data = self
        self.on_applying(cargs)
        if cargs.cancel:
            return
        _Events().trigger(FormatNamedEvent.STYLE_APPLYING, cargs)
        if cargs.cancel:
            return
        try:
            if self._is_valid_obj(obj):
                p = self._get_style_props(obj)
                # Could call p.setPropertyValue() here insetead of Props.set()
                # but by calling Props.set() events are triggered.
                mProps.Props.set(p, **self._get_properties())
                eargs = EventArgs.from_args(cargs)
                self.on_applied(eargs)
                _Events().trigger(FormatNamedEvent.STYLE_APPLIED, eargs)
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
    def prop_style_name(self, value: str):
        raise NotImplementedError
