"""
Module for Fill Transparency.

.. versionadded:: 0.9.0
"""

# region Imports
from __future__ import annotations
from typing import cast
import uno

from ooodev.format.inner.direct.write.frame.frame_type.position import Horizontal
from ooodev.format.inner.direct.write.frame.frame_type.position import Position as InnerPosition
from ooodev.format.inner.direct.write.frame.frame_type.position import Vertical
from ooodev.format.inner.modify.write.frame.frame_style_base_multi import FrameStyleBaseMulti
from ooodev.format.writer.style.frame.style_frame_kind import StyleFrameKind

# endregion Imports


class Position(FrameStyleBaseMulti):
    """
    Frame Style Type position.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        horizontal: Horizontal | None = None,
        vertical: Vertical | None = None,
        keep_boundaries: bool | None = None,
        mirror_even: bool | None = None,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> None:
        """
        Constructor

        Args:
            horizontal (Horizontal, optional): Specifies the Horizontal position options.
            vertical (Vertical, optional): Specifies the Vertical position options.
            keep_boundaries (bool, optional): Specifies keep inside text boundaries.
            mirror_even (bool, optional): Specifies mirror on even pages.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to.
                Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            None:
        """

        direct = InnerPosition(
            horizontal=horizontal, vertical=vertical, keep_boundaries=keep_boundaries, mirror_even=mirror_even
        )
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
    ) -> Position:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Default is
                Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            Position: ``Position`` instance from style properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerPosition.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerPosition:
        """Gets/Sets Inner Position instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerPosition, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerPosition) -> None:
        if not isinstance(value, InnerPosition):
            raise TypeError(f'Expected type of InnerPosition, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
