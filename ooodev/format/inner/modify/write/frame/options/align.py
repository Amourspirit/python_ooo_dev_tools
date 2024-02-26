# region Imports
from __future__ import annotations
from typing import cast

from ooodev.format.inner.direct.write.frame.options.align import Align as InnerAlign
from ooodev.format.inner.direct.write.frame.options.align import VertAdjustKind
from ooodev.format.inner.modify.write.frame.frame_style_base_multi import FrameStyleBaseMulti
from ooodev.format.writer.style.frame.style_frame_kind import StyleFrameKind

# endregion Imports


class Align(FrameStyleBaseMulti):
    """
    Frame Style Options Align.

    .. versionadded:: 0.9.0
    """

    # region Init
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
            adjust (VertAdjustKind): Specifies Vertical Adjustment. Default ``VertAdjustKind.TOP``
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to.
                Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            None:
        """

        direct = InnerAlign(adjust=adjust)
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    # endregion Init

    # region Static Methods
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
            doc (object): UNO Document Object.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to.
                Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            Align: ``Align`` instance from style properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerAlign.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    # endregion Static Methods

    # region Properties
    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleFrameKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerAlign:
        """Gets/Sets Inner Align instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerAlign, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerAlign) -> None:
        if not isinstance(value, InnerAlign):
            raise TypeError(f'Expected type of InnerAlign, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())

    # endregion Properties
