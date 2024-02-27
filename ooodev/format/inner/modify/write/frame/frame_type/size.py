"""
Module for Fill Transparency.

.. versionadded:: 0.9.0
"""

# region Imports
from __future__ import annotations
from typing import cast
import uno

from ooodev.format.inner.direct.write.image.image_type.size import AbsoluteSize
from ooodev.format.inner.direct.write.image.image_type.size import RelativeSize
from ooodev.format.inner.direct.write.frame.frame_type.size import Size as InnerSize
from ooodev.format.inner.modify.write.frame.frame_style_base_multi import FrameStyleBaseMulti
from ooodev.format.writer.style.frame.style_frame_kind import StyleFrameKind

# endregion Imports


class Size(FrameStyleBaseMulti):
    """
    Frame Style Type size.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        width: RelativeSize | AbsoluteSize | None = None,
        height: RelativeSize | AbsoluteSize | None = None,
        auto_width: bool = False,
        auto_height: bool = False,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> None:
        """
        Constructor

        Args:
            width (RelativeSize, AbsoluteSize, optional): width value.
            height (RelativeSize, AbsoluteSize, optional): height value.
            auto_width (bool, optional): Auto Size Width. Default ``False``.
            auto_height (bool, optional): Auto Size Height. Default ``False``.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to.
                Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            None:
        """

        direct = InnerSize(width=width, height=height, auto_width=auto_width, auto_height=auto_height)
        direct._prop_parent = self
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct)

    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> Size:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to.
                Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            Size: ``Size`` instance from style properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerSize.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct)
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleFrameKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerSize:
        """Gets/Sets Inner Size instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerSize, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerSize) -> None:
        if not isinstance(value, InnerSize):
            raise TypeError(f'Expected type of DirectSize, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
