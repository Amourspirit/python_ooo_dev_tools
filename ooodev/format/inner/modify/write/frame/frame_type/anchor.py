"""
Module for Fill Transparency.

.. versionadded:: 0.9.0
"""

# region Imports
from __future__ import annotations
from typing import cast
import uno

from ooodev.format.inner.modify.write.frame.frame_style_base_multi import FrameStyleBaseMulti
from ooodev.format.writer.style.frame.style_frame_kind import StyleFrameKind
from ooodev.format.inner.direct.write.frame.frame_type.anchor import AnchorKind
from ooodev.format.inner.direct.write.frame.frame_type.anchor import Anchor as InnerAnchor

# endregion Imports


class Anchor(FrameStyleBaseMulti):
    """
    Frame Style Type Anchor Position.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        anchor: AnchorKind = AnchorKind.AT_PARAGRAPH,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> None:
        """
        Constructor

        Args:
            anchor (AnchorKind): Specifies the anchor position. Default is ``AnchorKind.AT_PARAGRAPH``
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            None:
        """

        direct = InnerAnchor(anchor=anchor)
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
    ) -> Anchor:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            Anchor: ``Anchor`` instance from style properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerAnchor.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerAnchor:
        """Gets/Sets Inner Anchor instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerAnchor, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerAnchor) -> None:
        if not isinstance(value, InnerAnchor):
            raise TypeError(f'Expected type of InnerAnchor, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
