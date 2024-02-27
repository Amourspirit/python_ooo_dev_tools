# region Imports
from __future__ import annotations
from typing import cast
import uno

from ooodev.format.inner.modify.write.frame.frame_style_base_multi import FrameStyleBaseMulti
from ooodev.format.writer.style.frame.style_frame_kind import StyleFrameKind
from ooodev.format.inner.direct.write.frame.wrap.options import Options as InnerOptions

# endregion Imports


class Options(FrameStyleBaseMulti):
    """
    Frame Style Wrap Options.

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(
        self,
        *,
        first: bool | None = None,
        background: bool | None = None,
        contour: bool | None = None,
        outside: bool | None = None,
        overlap: bool | None = None,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> None:
        """
        Constructor

        Args:
            first (bool , optional): Specifies first paragraph.
            background (bool , optional): Specifies in background.
            contour (bool , optional): Specifies contour.
            outside (bool , optional): Specifies contour outside only. ``contour`` must be ``True`` for this parameter to be effective.
            overlap (bool , optional): Specifies allow overlap.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to.
                Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            None:
        """

        direct = InnerOptions(first=first, background=background, contour=contour, outside=outside, overlap=overlap)
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
    ) -> Options:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to.
                Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            Options: ``Options`` instance from style properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerOptions.from_obj(inst.get_style_props(doc))
        direct._prop_parent = inst
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
    def prop_inner(self) -> InnerOptions:
        """Gets/Sets Inner Options instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerOptions, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerOptions) -> None:
        if not isinstance(value, InnerOptions):
            raise TypeError(f'Expected type of InnerOptions, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())

    # endregion Properties
