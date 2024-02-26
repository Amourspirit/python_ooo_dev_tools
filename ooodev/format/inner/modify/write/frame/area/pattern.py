# region Imports
from __future__ import annotations
from typing import cast
import uno
from com.sun.star.awt import XBitmap

from ooodev.format.inner.preset.preset_pattern import PresetPatternKind
from ooodev.format.inner.modify.write.frame.frame_style_base_multi import FrameStyleBaseMulti
from ooodev.format.writer.style.frame.style_frame_kind import StyleFrameKind
from ooodev.format.inner.direct.write.fill.area.pattern import Pattern as InnerPattern

# endregion Imports


class Pattern(FrameStyleBaseMulti):
    """
    Frame Style Area Pattern.

    .. versionadded:: 0.9.0
    """

    # region Init
    def __init__(
        self,
        *,
        bitmap: XBitmap | None = None,
        name: str = "",
        tile: bool = True,
        stretch: bool = False,
        auto_name: bool = False,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> None:
        """
        Constructor

        Args:
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table then this property is required.
            name (str, optional): Specifies the name of the pattern. This is also the name that is used to store bitmap in LibreOffice Bitmap Table.
            tile (bool, optional): Specified if bitmap is tiled. Defaults to ``True``.
            stretch (bool, optional): Specifies if bitmap is stretched. Defaults to ``False``.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            None:

        Note:
            If ``auto_name`` is ``False`` then a bitmap for a given name is only required the first call.
            All subsequent call of the same name will retrieve the bitmap form the LibreOffice Bitmap Table.
        """
        # pylint: disable=protected-access
        # pylint: disable=unexpected-keyword-arg
        direct = InnerPattern(
            bitmap=bitmap,
            name=name,
            tile=tile,
            stretch=stretch,
            auto_name=auto_name,
            _cattribs=self._get_inner_cattribs(),
        )
        direct._prop_parent = self
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family
        self._set_style("direct", direct, *direct.get_attrs())  # type: ignore

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
    ) -> Pattern:
        """
        Gets instance from Document.

        Args:
            doc (object): UNO Document Object.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            Pattern: ``Pattern`` instance from style properties.
        """
        # pylint: disable=protected-access
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerPattern.from_obj(obj=inst.get_style_props(doc), _cattribs=inst._get_inner_cattribs())
        direct._prop_parent = inst
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @classmethod
    def from_preset(
        cls,
        preset: PresetPatternKind,
        style_name: StyleFrameKind | str = StyleFrameKind.FRAME,
        style_family: str = "FrameStyles",
    ) -> Pattern:
        """
        Gets an instance from a preset.

        Args:
            preset (PresetPatternKind): Preset.
            style_name (StyleFrameKind, str, optional): Specifies the Frame Style that instance applies to. Default is Default Frame Style.
            style_family (str, optional): Style family. Default ``FrameStyles``.

        Returns:
            Pattern: ``Pattern`` instance from preset.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerPattern.from_preset(preset=preset, _cattribs=inst._get_inner_cattribs())
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
    def prop_inner(self) -> InnerPattern:
        """Gets Inner Pattern instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerPattern, self._get_style_inst("direct"))
        return self._direct_inner

    # endregion Properties
