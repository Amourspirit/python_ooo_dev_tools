"""
Module for Fill Transparency.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import cast
import uno

from ..frame_style_base_multi import FrameStyleBaseMulti
from ....writer.style.frame.style_frame_kind import StyleFrameKind as StyleFrameKind
from ....direct.frame.options.align import Align as DirectAlign, VertAdjustKind as VertAdjustKind


class Align(FrameStyleBaseMulti):
    """
    Frame Style Options Align.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        adjust: VertAdjustKind = VertAdjustKind.TOP,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> None:
        """
        Constructor

        Args:
            adjust (VertAdjustKindl): Specifies Verticial Adjustment. Default ``VertAdjustKind.TOP``
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Deftult is Default Frame Style.
            style_family (str, optional): Style family. Defatult ``FrameStyles``.

        Returns:
            None:
        """

        direct = DirectAlign(adjust=adjust)
        direct._prop_parent = self
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> Align:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Documnet Object.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Deftult is Default Frame Style.
            style_family (str, optional): Style family. Defatult ``FrameStyles``.

        Returns:
            Align: ``Align`` instance from style properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = DirectAlign.from_obj(inst.get_style_props(doc))
        direct._prop_parent = inst
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleFrameKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> DirectAlign:
        """Gets Inner Align instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(DirectAlign, self._get_style_inst("direct"))
        return self._direct_inner