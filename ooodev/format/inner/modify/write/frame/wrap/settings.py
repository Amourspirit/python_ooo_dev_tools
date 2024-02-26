# region Imports
from __future__ import annotations
from typing import cast
import uno
from ooo.dyn.text.wrap_text_mode import WrapTextMode

from ooodev.format.inner.modify.write.frame.frame_style_base_multi import FrameStyleBaseMulti
from ooodev.format.writer.style.frame.style_frame_kind import StyleFrameKind
from ooodev.format.inner.direct.write.frame.wrap.settings import Settings as InnerSettings

# endregion Imports


class Settings(FrameStyleBaseMulti):
    """
    Frame Style Wrap Settings.

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        mode: WrapTextMode = WrapTextMode.PARALLEL,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> None:
        """
        Constructor

        Args:
            mode (WrapTextMode): Specifies Wrap mode. Default ``WrapTextMode.PARALLEL``
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to.
                Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            None:
        """

        direct = InnerSettings(mode=mode)
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
    ) -> Settings:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to.
                Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            Settings: ``Settings`` instance from style properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerSettings.from_obj(inst.get_style_props(doc))
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
    def prop_inner(self) -> InnerSettings:
        """Gets/Sets Inner Settings instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerSettings, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerSettings) -> None:
        if not isinstance(value, InnerSettings):
            raise TypeError(f'Expected type of InnerSettings, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
