# region Imports
from __future__ import annotations
from typing import cast

from ..frame_style_base_multi import FrameStyleBaseMulti
from ooodev.format.writer.style.frame.style_frame_kind import StyleFrameKind as StyleFrameKind
from ooodev.format.inner.direct.write.frame.options.protect import Protect as InnerProtect

# endregion Imports


class Protect(FrameStyleBaseMulti):
    """
    Frame Style Options Protect.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        size: bool | None = None,
        position: bool | None = None,
        content: bool | None = None,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> None:
        """
        Constructor

        Args:
            size (bool, optional): Specifies size protection.
            position (bool, optional): Specifies position protection.
            content (bool, optional): Specifies content protection.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to.
                Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            None:
        """

        direct = InnerProtect(size=size, position=position, content=content)
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
    ) -> Protect:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to.
                Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            Protect: ``Protect`` instance from style properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerProtect.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerProtect:
        """Gets/Sets Inner Protect instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerProtect, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerProtect) -> None:
        if not isinstance(value, InnerProtect):
            raise TypeError(f'Expected type of InnerProtect, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
