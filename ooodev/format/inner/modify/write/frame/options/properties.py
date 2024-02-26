from __future__ import annotations
from typing import cast

from ooodev.format.inner.direct.write.frame.options.properties import Properties as InnerProperties
from ooodev.format.inner.direct.write.frame.options.properties import TextDirectionKind
from ooodev.format.inner.modify.write.frame.frame_style_base_multi import FrameStyleBaseMulti
from ooodev.format.writer.style.frame.style_frame_kind import StyleFrameKind


class Properties(FrameStyleBaseMulti):
    """
    Frame Style Options Properties.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        editable: bool | None = None,
        printable: bool | None = None,
        txt_direction: TextDirectionKind | None = None,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> None:
        """
        Constructor

        Args:
            editable (bool, optional): Specifies if Frame is editable in read-only document.
            printable (bool, optional): Specifies if Frame can be printed.
            txt_direction (TextDirectionKind, optional): Specifies text direction.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to.
                Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            None:
        """

        direct = InnerProperties(editable=editable, printable=printable, txt_direction=txt_direction)
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
    ) -> Properties:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to.
                Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            Properties: ``Properties`` instance from style properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerProperties.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerProperties:
        """Gets/Sets Inner Properties instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerProperties, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerProperties) -> None:
        if not isinstance(value, InnerProperties):
            raise TypeError(f'Expected type of InnerProperties, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
