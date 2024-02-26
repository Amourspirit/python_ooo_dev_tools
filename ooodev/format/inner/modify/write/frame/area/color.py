# region Imports
from __future__ import annotations
from typing import cast
import uno

from ooodev.format.inner.modify.write.frame.frame_style_base_multi import FrameStyleBaseMulti
from ooodev.format.writer.style.frame.style_frame_kind import StyleFrameKind
from ooodev.format.inner.direct.write.fill.area.fill_color import FillColor as InnerColor
from ooodev.utils import color as mColor
from ooodev.utils.color import StandardColor

# endregion Imports


class Color(FrameStyleBaseMulti):
    """
    Frame Style Area Color.

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(
        self,
        *,
        color: mColor.Color = StandardColor.AUTO_COLOR,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> None:
        """
        Constructor

        Args:
            color (:py:data:`~.utils.color.Color`, optional): Color.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to.
                Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            None:
        """

        direct = InnerColor(color=color, _cattribs=self._get_inner_cattribs())
        direct._prop_parent = self
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())

    # endregion Init

    # region Internal Methods
    def _get_inner_cattribs(self) -> dict:
        return {"_supported_services_values": self._supported_services(), "_format_kind_prop": self.prop_format_kind}

    # endregion Internal Methods

    # region Static Methods
    @classmethod
    def from_style(
        cls,
        doc: object,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> Color:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            Color: ``Color`` instance from style properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerColor.from_obj(obj=inst.get_style_props(doc), _cattribs=inst._get_inner_cattribs())
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
    def prop_inner(self) -> InnerColor:
        """Gets/Sets Inner Color instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerColor, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerColor) -> None:
        if not isinstance(value, InnerColor):
            raise TypeError(f'Expected type of InnerColor, got "{type(value).__name__}"')
        inst = value.__class__(_cattribs=self._get_inner_cattribs())
        for prop in value._props:
            if prop:
                inst._set(prop, value._get(prop))
        self._del_attribs("_direct_inner")
        self._set_style("direct", inst, *inst.get_attrs())

    # endregion Properties
