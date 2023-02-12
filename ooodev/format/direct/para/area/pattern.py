from __future__ import annotations
from typing import Any, Tuple, cast, overload, Type, TypeVar

import uno

from .....events.args.key_val_cancel_args import KeyValCancelArgs
from .....exceptions import ex as mEx
from .....utils import lo as mLo
from .....utils import props as mProps
from ....kind.format_kind import FormatKind
from ....preset.preset_pattern import PresetPatternKind as PresetPatternKind
from ....style_base import StyleMulti
from ...fill.area.pattern import Pattern as FillPattern

from com.sun.star.awt import XBitmap

from ooo.dyn.style.graphic_location import GraphicLocation

_TPattern = TypeVar(name="_TPattern", bound="Pattern")


class Pattern(StyleMulti):
    def __init__(
        self,
        *,
        bitmap: XBitmap | None = None,
        name: str = "",
        tile: bool = True,
        stretch: bool = False,
        auto_name: bool = False,
    ) -> None:
        """
        Constructor

        Args:
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table then this property is requied.
            name (str, optional): Specifies the name of the pattern. This is also the name that is used to store bitmap in LibreOffice Bitmap Table.
            tile (bool, optional): Specified if bitmap is tiled. Defaults to ``True``.
            stretch (bool, optional): Specifies if bitmap is stretched. Defaults to ``False``.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.

        Returns:
            None:

        Note:
            If ``auto_name`` is ``False`` then a bitmap for a given name is only required the first call.
            All subsequent call of the same name will retreive the bitmap form the LibreOffice Bitmap Table.
        """

        fp = FillPattern(bitmap=bitmap, name=name, tile=tile, stretch=stretch, auto_name=auto_name)
        if stretch:
            loc = GraphicLocation.AREA
        else:
            loc = GraphicLocation.TILED
        init_vars = {
            "ParaBackColor": -1,
            "ParaBackGraphicLocation": loc,
            "ParaBackTransparent": True,
        }
        bmap = fp._get("FillBitmap")
        if not bmap is None:
            init_vars["ParaBackGraphic"] = bmap

        super().__init__(**init_vars)
        self._set_style("fill_props", fp, *fp.get_attrs())

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        return (
            "com.sun.star.style.ParagraphProperties",
            "com.sun.star.text.TextContent",
            "com.sun.star.beans.PropertySet",
            "com.sun.star.style.ParagraphStyle",
        )

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"FillPattern.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO object.

        Returns:
            None:
        """
        fp = cast(FillPattern, self._get_style("fill_props")[0])
        if not fp._has("FillBitmap"):
            mLo.Lo.print("Pattern.apply(): There is nothing to apply.")
            return
        super().apply(obj, **kwargs)

    def on_property_restore_setting(self, event_args: KeyValCancelArgs) -> None:
        defaults = ("ParaBackColor", "ParaBackGraphicLocation", "ParaBackTransparent")
        if event_args.key in defaults:
            default = mProps.Props.get_default(event_args.event_data, event_args.key)
            if default == event_args.value:
                event_args.default = True
        elif event_args.key == "ParaBackGraphic":
            if not event_args.value:
                event_args.cancel = True

        return super().on_property_restore_setting(event_args)

    # endregion Overrides

    # region Static Methods
    @classmethod
    def from_preset(cls: Type[_TPattern], preset: PresetPatternKind) -> _TPattern:
        """
        Gets an instance from a preset

        Args:
            preset (PatternKind): Preset

        Returns:
            Pattern: Instance from preset.
        """
        fp = FillPattern.from_preset(preset)
        bmap = fp._get("FillBitmap")
        name = str(preset)
        inst = cast(Pattern, super(Pattern, cls).__new__(cls))
        inst.__init__(bitmap=bmap, name=name, tile=True, stretch=False, auto_name=False)
        return inst

    @classmethod
    def from_obj(cls: Type[_TPattern], obj: object) -> _TPattern:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Pattern: ``Pattern`` instance that represents ``obj`` fill pattern.
        """
        fp = FillPattern.from_obj(obj)
        bmap = fp._get("FillBitmap")
        name = fp._get("FillBitmapName")
        inst = cast(Pattern, super(Pattern, cls).__new__(cls))
        inst.__init__(bitmap=bmap, name=name, tile=True, stretch=False, auto_name=False)
        return inst

    # endregion Static Methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.TXT_CONTENT | FormatKind.PARA

    @property
    def prop_inner(self) -> FillPattern:
        """Gets Fill Pattern instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(FillPattern, self._get_style_inst("fill_props"))
        return self._direct_inner

    # endregion Properties
