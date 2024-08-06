from __future__ import annotations
from typing import Any, Tuple, cast, overload, Type, TypeVar
import uno
from com.sun.star.awt import XBitmap
from ooo.dyn.style.graphic_location import GraphicLocation

from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.preset.preset_pattern import PresetPatternKind
from ooodev.format.inner.style_base import StyleMulti
from ooodev.format.inner.direct.write.fill.area.pattern import Pattern as InnerPattern


_TPattern = TypeVar("_TPattern", bound="Pattern")


class Pattern(StyleMulti):
    """
    Class for Para Fill Pattern.

    .. seealso::

        - :ref:`help_writer_format_direct_para_area_pattern`

    .. versionadded:: 0.9.0
    """

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
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table then
                this property is required.
            name (str, optional): Specifies the name of the pattern. This is also the name that is used to store
                bitmap in LibreOffice Bitmap Table.
            tile (bool, optional): Specified if bitmap is tiled. Defaults to ``True``.
            stretch (bool, optional): Specifies if bitmap is stretched. Defaults to ``False``.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.

        Returns:
            None:

        Note:
            If ``auto_name`` is ``False`` then a bitmap for a given name is only required the first call.
            All subsequent call of the same name will retrieve the bitmap form the LibreOffice Bitmap Table.

        See Also:

            - :ref:`help_writer_format_direct_para_area_pattern`
        """

        fp = InnerPattern(bitmap=bitmap, name=name, tile=tile, stretch=stretch, auto_name=auto_name)
        loc = GraphicLocation.AREA if stretch else GraphicLocation.TILED
        init_vars = {
            "ParaBackColor": -1,
            "ParaBackGraphicLocation": loc,
            "ParaBackTransparent": True,
        }
        b_map = fp._get("FillBitmap")
        if b_map is not None:
            init_vars["ParaBackGraphic"] = b_map

        super().__init__(**init_vars)
        fp._prop_parent = self
        self._set_style("fill_props", fp, *fp.get_attrs())  # type: ignore

    # region Overrides

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.ParagraphProperties",
                "com.sun.star.text.TextContent",
                "com.sun.star.beans.PropertySet",
                "com.sun.star.style.ParagraphStyle",
            )
        return self._supported_services_values

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print("FillPattern.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, **kwargs) -> None: ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO object.

        Returns:
            None:
        """
        fp = cast(InnerPattern, self._get_style_inst("fill_props"))
        if not fp._has("FillBitmap"):
            mLo.Lo.print("Pattern.apply(): There is nothing to apply.")
            return
        super().apply(obj, **kwargs)

    def on_property_restore_setting(self, source: Any, event_args: KeyValCancelArgs) -> None:
        defaults = ("ParaBackColor", "ParaBackGraphicLocation", "ParaBackTransparent")
        if event_args.key in defaults:
            default = mProps.Props.get_default(event_args.event_data, event_args.key)
            if default == event_args.value:
                event_args.default = True
        elif event_args.key == "ParaBackGraphic":
            if not event_args.value:
                event_args.cancel = True

        return super().on_property_restore_setting(source, event_args)

    # endregion Overrides

    # region Static Methods

    # region from_preset()
    @overload
    @classmethod
    def from_preset(cls: Type[_TPattern], preset: PresetPatternKind) -> _TPattern: ...

    @overload
    @classmethod
    def from_preset(cls: Type[_TPattern], preset: PresetPatternKind, **kwargs) -> _TPattern: ...

    @classmethod
    def from_preset(cls: Type[_TPattern], preset: PresetPatternKind, **kwargs) -> _TPattern:
        """
        Gets an instance from a preset

        Args:
            preset (PresetPatternKind): Preset

        Returns:
            Pattern: Instance from preset.
        """
        # pylint: disable=protected-access
        fp = InnerPattern.from_preset(preset)
        bitmap = fp._get("FillBitmap")
        name = str(preset)
        return cls(bitmap=bitmap, name=name, tile=True, stretch=False, auto_name=False, **kwargs)

    # endregion from_preset()

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TPattern], obj: Any) -> _TPattern: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TPattern], obj: Any, **kwargs) -> _TPattern: ...

    @classmethod
    def from_obj(cls: Type[_TPattern], obj: Any, **kwargs) -> _TPattern:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Pattern: ``Pattern`` instance that represents ``obj`` fill pattern.
        """
        # pylint: disable=protected-access
        fp = InnerPattern.from_obj(obj)
        bitmap = fp._get("FillBitmap")
        name = fp._get("FillBitmapName")
        result = cls(bitmap=bitmap, name=name, tile=True, stretch=False, auto_name=False, **kwargs)
        result.set_update_obj(obj)
        return result

    # endregion from_obj()

    # endregion Static Methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.TXT_CONTENT | FormatKind.PARA
        return self._format_kind_prop

    @property
    def prop_inner(self) -> InnerPattern:
        """Gets Fill Pattern instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerPattern, self._get_style_inst("fill_props"))
        return self._direct_inner

    # endregion Properties
