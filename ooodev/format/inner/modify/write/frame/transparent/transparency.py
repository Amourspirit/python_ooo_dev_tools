# region Imports
from __future__ import annotations
from typing import cast
import uno

from ooodev.utils.data_type.intensity import Intensity
from ooodev.format.inner.modify.write.frame.frame_style_base_multi import FrameStyleBaseMulti
from ooodev.format.writer.style.frame.style_frame_kind import StyleFrameKind
from ooodev.format.inner.direct.write.fill.transparent.transparency import Transparency as InnerTransparency

# endregion Imports


class Transparency(FrameStyleBaseMulti):
    """
    Frame Style Transparency Transparency.

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(
        self,
        *,
        value: Intensity | int = 0,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> None:
        """
        Constructor

        Args:
            value (Intensity, int, optional): Specifies the transparency value from ``0`` to ``100``.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to.
                Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            None:
        """

        direct = InnerTransparency(value=value, _cattribs=self._get_inner_cattribs())
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
    ) -> Transparency:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to.
                Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            Transparency: ``Transparency`` instance from style properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerTransparency.from_obj(obj=inst.get_style_props(doc), _cattribs=inst._get_inner_cattribs())
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    # endregion Static Methods

    # region internal methods
    def _get_inner_cattribs(self) -> dict:
        return {"_format_kind_prop": self.prop_format_kind, "_supported_services_values": self._supported_services()}

    # endregion internal methods

    # region Properties
    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleFrameKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerTransparency:
        """Gets/Sets Inner Transparency instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerTransparency, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerTransparency) -> None:
        if not isinstance(value, InnerTransparency):
            raise TypeError(f'Expected type of InnerTransparency, got "{type(value).__name__}"')
        direct = value.__class__(value=value.prop_value, _cattribs=self._get_inner_cattribs())
        self._del_attribs("_direct_inner")
        self._set_style("direct", direct, *direct.get_attrs())

    # endregion Properties
