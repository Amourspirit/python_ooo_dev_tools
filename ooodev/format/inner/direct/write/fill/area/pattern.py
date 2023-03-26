"""
Module for Fill Properties Fill Pattern.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Any, Tuple, overload, Type, TypeVar

from com.sun.star.awt import XBitmap

from ooo.dyn.drawing.fill_style import FillStyle as FillStyle

from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.preset import preset_pattern as mPattern
from ooodev.format.inner.preset.preset_pattern import PresetPatternKind as PresetPatternKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.format.inner.common.props.area_pattern_props import AreaPatternProps


# https://github.com/LibreOffice/core/blob/6379414ca34527fbe69df2035d49d651655317cd/vcl/source/filter/ipict/ipict.cxx#L92

_TPattern = TypeVar(name="_TPattern", bound="Pattern")


class Pattern(StyleBase):
    """
    Class for Fill Properties Fill Pattern.

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
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table then this property is required.
            name (str, optional): Specifies the name of the pattern. This is also the name that is used to store bitmap in LibreOffice Bitmap Table.
            tile (bool, optional): Specified if bitmap is tiled. Defaults to ``True``.
            stretch (bool, optional): Specifies if bitmap is stretched. Defaults to ``False``.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.

        Returns:
            None:

        Note:
            If ``auto_name`` is ``False`` then a bitmap for a given name is only required the first call.
            All subsequent call of the same name will retrieve the bitmap form the LibreOffice Bitmap Table.
        """

        init_vals = {}
        # if bitmap or name is passed in then get the bitmap
        init_vals[self._props.tile] = tile
        init_vals[self._props.stretch] = stretch
        bmap = None
        try:
            bmap = self._get_bitmap(bitmap, name, auto_name)
        except Exception:
            pass
        if not bmap is None:
            init_vals[self._props.bitmap] = bmap
            init_vals[self._props.name] = self._name
            init_vals[self._props.style] = FillStyle.BITMAP

        super().__init__(**init_vals)

    # region Internal Methods
    def _get_bitmap(self, bitmap: XBitmap | None, name: str, auto_name: bool) -> XBitmap:
        # if the name passed in already exist in the Bitmap Table then it is returned.
        # Otherwise the bitmap is added to the Bitmap Table and then returned.
        # after bitmap is added to table all other subsequent call of this name will return
        # that bitmap from the Table. With the exception of auto_name which will force a new entry
        # into the Table each time.
        if not name:
            auto_name = True
            name = "FILL PATTERN "
        nc = self._container_get_inst()
        if auto_name:
            self._name = self._container_get_unique_el_name(name, nc)
        else:
            self._name = name
        bmap = self._container_get_value(self._name, nc)  # raises value error if name is empty
        if not bmap is None:
            return bmap
        if bitmap is None:
            raise ValueError(f'No bitmap could be found in container for "{name}". In this case a bitmap is required.')
        self._container_add_value(name=self._name, obj=bitmap, allow_update=False, nc=nc)
        return self._container_get_value(self._name, nc)

    # endregion Internal Methods

    # region Overrides
    # region copy()
    @overload
    def copy(self: _TPattern) -> _TPattern:
        ...

    @overload
    def copy(self: _TPattern, **kwargs) -> _TPattern:
        ...

    def copy(self: _TPattern, **kwargs) -> _TPattern:
        """Gets a copy of instance as a new instance"""
        cp = super().copy(**kwargs)
        cp._name = self._name
        return cp

    # endregion copy()

    def _container_get_service_name(self) -> str:
        return "com.sun.star.drawing.BitmapTable"

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.beans.PropertySet",
                "com.sun.star.drawing.FillProperties",
                "com.sun.star.style.PageStyle",
                "com.sun.star.style.ParagraphStyle",
                "com.sun.star.text.BaseFrame",
                "com.sun.star.text.TextContent",
                "com.sun.star.text.TextEmbeddedObject",
                "com.sun.star.text.TextFrame",
                "com.sun.star.text.TextGraphicObject",
            )
        return self._supported_services_values

    # region apply()
    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO object that has supports ``com.sun.star.style.FillProperties`` or ``com.sun.star.beans.PropertySet`` service.

        Returns:
            None:
        """
        if not self._has(self._props.bitmap):
            mLo.Lo.print("Pattern.apply(): There is nothing to apply.")
            return
        super().apply(obj, **kwargs)

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"Pattern.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()

    def on_property_restore_setting(self, source: Any, event_args: KeyValCancelArgs) -> None:
        if event_args.key == self._props.bitmap:
            if event_args.value is None:
                event_args.default = True
        elif event_args.key == self._props.name:
            if not event_args.value:
                event_args.default = True
        return super().on_property_restore_setting(source, event_args)

    # endregion Overrides

    # region Static Methods
    # region from_preset()
    @overload
    @classmethod
    def from_preset(cls: Type[_TPattern], preset: PresetPatternKind) -> _TPattern:
        ...

    @overload
    @classmethod
    def from_preset(cls: Type[_TPattern], preset: PresetPatternKind, **kwargs) -> _TPattern:
        ...

    @classmethod
    def from_preset(cls: Type[_TPattern], preset: PresetPatternKind, **kwargs) -> _TPattern:
        """
        Gets an instance from a preset.

        Args:
            preset (PatternKind): Preset.

        Returns:
            Pattern: ``Pattern`` instance from preset.
        """
        name = str(preset)
        nu = cls(**kwargs)

        nc = nu._container_get_inst()
        bmap = nu._container_get_value(name, nc)
        if bmap is None:
            bmap = mPattern.get_prest_bitmap(preset)
        return cls(bitmap=bmap, name=name, tile=True, stretch=False, auto_name=False, **kwargs)

    # endregion from_preset()

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TPattern], obj: object) -> _TPattern:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TPattern], obj: object, **kwargs) -> _TPattern:
        ...

    @classmethod
    def from_obj(cls: Type[_TPattern], obj: object, **kwargs) -> _TPattern:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Pattern: ``Pattern`` instance that represents ``obj`` fill pattern.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        def set_prop(key: str, fp: Pattern):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if not val is None:
                fp._set(key, val)

        name = mProps.Props.get(obj, inst._props.name)
        inst._name = name
        inst._set(inst._props.name, name)
        set_prop(inst._props.tile, inst)
        set_prop(inst._props.stretch, inst)
        set_prop(inst._props.bitmap, inst)
        set_prop(inst._props.style, inst)
        return inst

    # endregion from_obj()
    # endregion Static Methods

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.TXT_CONTENT | FormatKind.FILL
        return self._format_kind_prop

    @property
    def prop_tile(self) -> bool:
        """Gets sets if fill image is tiled"""
        return self._get(self._props.tile)

    @prop_tile.setter
    def prop_tile(self, value: bool):
        self._set(self._props.tile, value)

    @property
    def prop_stretch(self) -> bool:
        """Gets sets if fill image is stretched"""
        return self._get(self._props.stretch)

    @prop_stretch.setter
    def prop_stretch(self, value: bool):
        self._set(self._props.stretch, value)

    @property
    def _props(self) -> AreaPatternProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = AreaPatternProps(
                style="FillStyle",
                name="FillBitmapName",
                tile="FillBitmapTile",
                stretch="FillBitmapStretch",
                bitmap="FillBitmap",
            )
        return self._props_internal_attributes

    # endregion Properties
