"""
Module for managing character Font position.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Any, Tuple, Type, cast, overload, TypeVar
from enum import Enum

from .....events.args.cancel_event_args import CancelEventArgs
from .....events.args.key_val_cancel_args import KeyValCancelArgs
from .....exceptions import ex as mEx
from .....meta.static_prop import static_prop
from .....utils import lo as mLo
from .....utils import props as mProps
from .....utils.data_type.angle import Angle
from .....utils.data_type.intensity import Intensity as Intensity
from .....utils.unit_convert import UnitConvert
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase

_TFontPosition = TypeVar(name="_TFontPosition", bound="FontPosition")


class CharSpacingKind(Enum):
    """
    Character Spacing
    """

    VERY_TIGHT = -3.0
    TIGHT = -1.5
    NORMAL = 0.0
    LOOSE = 3.0
    VERY_LOOSE = 6.0

    def __float__(self) -> float:
        return self.value


class FontScriptKind(Enum):
    """Font Script Type"""

    NORMAL = 0
    """Normal"""
    SUPERSCRIPT = 14_000
    """Superscript"""
    SUBSCRIPT = -14_000
    """Subscript"""

    def __int__(self) -> int:
        return self.value


class FontPosition(StyleBase):
    """
    Character Font Position

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together font properties.

    Many properties can be chained together.

    .. versionadded:: 0.9.0
    """

    _DEFAULT_SUPER_SUB_HEIGHT = 58

    def __init__(
        self,
        *,
        script_kind: FontScriptKind | None = None,
        raise_lower: int | Intensity | None = None,
        rel_size: int | None = None,
        rotation: int | Angle | None = None,
        scale: int | None = None,
        fit: bool | None = None,
        spacing: CharSpacingKind | float | None = None,
        pair: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            script_kind (FontScriptKind, optional): Specifies Superscript/Subscript option.
            raise_lower (int, Intensity, optional): Specifies raise or Lower.
            rel_size (int, optional): Specifies realitive Font Size. Set this value to ``0`` for automatic.
            rotation (int, Angle, optional): Specifies the rotation of a character in degrees. Depending on the implementation only certain values may be allowed.
            scale (int, optional): Specifies scale width.
            fit (bool, optional): Specifies if rotation is fit to line.
            spacing (float, optional): Specifies character spacing in point units.
            pair (bool, optional): Specifies pair kerning.
        """

        super().__init__()
        # superscript and subscript use the same internal properties,CharEscapementHeight, CharEscapement
        if not script_kind is None:
            self.prop_script_kind = script_kind
        if not raise_lower is None:
            self.prop_raise_lower = raise_lower
        if not rel_size is None:
            self.prop_rel_size = rel_size
        if not rotation is None:
            self.prop_rotation = rotation
        if not scale is None:
            self.prop_scale = scale
        if not fit is None:
            self.prop_fit = fit
        if not spacing is None:
            self.prop_spacing = spacing
        if not pair is None:
            self.prop_pair = pair

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        return ("com.sun.star.style.CharacterProperties", "com.sun.star.style.CharacterStyle")

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_setting(event)

    # region apply()
    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO object that has supports ``com.sun.star.style.CharacterProperties`` service.

        Returns:
            None:
        """
        super().apply(obj, **kwargs)

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"FontPosition.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()

    @classmethod
    def from_obj(cls: Type[_TFontPosition], obj: object) -> _TFontPosition:
        """
        Gets instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.CharacterProperties`` service.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            FontPosition: ``FontPosition`` instance that represents ``obj`` font postiion.
        """
        inst = super(FontPosition, cls).__new__(cls)
        inst.__init__()
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        def set_prop(key: str, fp: FontPosition):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if not val is None:
                fp._set(key, val)

        set_prop("CharEscapement", inst)
        set_prop("CharEscapementHeight", inst)
        set_prop("CharRotation", inst)
        set_prop("CharScaleWidth", inst)
        set_prop("CharRotationIsFitToLine", inst)
        set_prop("CharKerning", inst)
        set_prop("CharAutoKerning", inst)
        return inst

    def on_property_setting(self, event_args: KeyValCancelArgs) -> None:
        """
        Triggers for each property that is set

        Args:
            event_args (KeyValueCancelArgs): Event Args
        """
        # if position is normal then defaults should be set
        if event_args.key == "CharEscapementHeight":
            if self.prop_script_kind == FontScriptKind.NORMAL:
                event_args.value == 100
        if event_args.key == "CharEscapement":
            if self.prop_script_kind == FontScriptKind.NORMAL:
                event_args.value == 0
        super().on_property_setting(event_args)

    # endregion methods

    # region Format Methods
    def fmt_scrip_kind(self: _TFontPosition, value: FontScriptKind | None = None) -> _TFontPosition:
        """
        Get copy of instance with superscript/subscript set or removed.

        Args:
            value (FontScriptKind, optional): font script kind.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontPosition: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_script_kind = value
        return ft

    def fmt_raise_lower(self: _TFontPosition, value: int | Intensity | None = None) -> _TFontPosition:
        """
        Get copy of instance with raise/lower set or removed.

        Args:
            value (int, Intensity, optional): Raise or Lower value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontPosition: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_rel_size = value
        return ft

    def fmt_rel_size(self: _TFontPosition, value: int | None = None) -> _TFontPosition:
        """
        Get copy of instance with relative size set or removed.

        Args:
            value (int, Intensity, optional): relative size value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontPosition: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_raise_lower = value
        return ft

    def fmt_rotation(self: _TFontPosition, value: int | Angle | None = None) -> _TFontPosition:
        """
        Get copy of instance with rotation set or removed.

        Args:
            value (int, Angle, optional): The rotation of a character in degrees. Depending on the implementation only certain values may be allowed.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontPosition: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_rotation = value
        return ft

    def fmt_scale(self: _TFontPosition, value: int | None = None) -> _TFontPosition:
        """
        Get copy of instance with scale width set or removed.

        Args:
            value (int, Intensity, optional): scale width value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontPosition: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_scale = value
        return ft

    def fmt_fit(self: _TFontPosition, value: bool | None = None) -> _TFontPosition:
        """
        Get copy of instance with rotation fit set or removed.

        Args:
            value (int, Intensity, optional): Rotation fit value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontPosition: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_fit = value
        return ft

    def fmt_spacing(self: _TFontPosition, value: float | None = None) -> _TFontPosition:
        """
        Get copy of instance with spacing set or removed.

        Args:
            value (str, optional): The character spacing in point units.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontPosition: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_spacing = value
        return ft

    def fmt_pair(self: _TFontPosition, value: bool | None = None) -> _TFontPosition:
        """
        Get copy of instance with pair kerning set or removed.

        Args:
            value (int, Intensity, optional): Pair kerning value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontPosition: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_pair = value
        return ft

    # endregion Format Methods

    # region Style Properties
    @property
    def script_kind_normal(self: _TFontPosition) -> _TFontPosition:
        """Gets copy of instance set to Position Normal"""
        ft = self.copy()
        ft.prop_script_kind = FontScriptKind.NORMAL
        return ft

    @property
    def script_kind_superscript(self: _TFontPosition) -> _TFontPosition:
        """Gets copy of instance set to Position Superscript"""
        ft = self.copy()
        ft.prop_script_kind = FontScriptKind.SUPERSCRIPT
        return ft

    @property
    def script_kind_subscript(self: _TFontPosition) -> _TFontPosition:
        """Gets copy of instance set to Position Subscript"""
        ft = self.copy()
        ft.prop_script_kind = FontScriptKind.SUBSCRIPT
        return ft

    @property
    def raise_lower_auto(self: _TFontPosition) -> _TFontPosition:
        """Gets copy of instance with raise/lower set to automatic"""
        ft = self.copy()
        ft.prop_raise_lower = 0
        return ft

    @property
    def rotation_none(self: _TFontPosition) -> _TFontPosition:
        """Gets copy of instance with rotation set to ``0``"""
        ft = self.copy()
        ft.prop_rotation = 0
        return ft

    @property
    def rotation_90(self: _TFontPosition) -> _TFontPosition:
        """Gets copy of instance with rotation set to ``90``"""
        ft = self.copy()
        ft.prop_rotation = 90
        return ft

    @property
    def rotation_270(self: _TFontPosition) -> _TFontPosition:
        """Gets copy of instance with rotation set to ``270``"""
        ft = self.copy()
        ft.prop_rotation = 270
        return ft

    @property
    def fit(self: _TFontPosition) -> _TFontPosition:
        """Gets copy of instance with rotation fit to line set to ``True``"""
        ft = self.copy()
        ft.prop_fit = True
        return ft

    @property
    def spacing_very_tight(self) -> _TFontPosition:
        """Gets copy of instance with spacing set to very tight value"""
        ft = self.copy()
        ft.prop_spacing = CharSpacingKind.VERY_TIGHT
        return ft

    @property
    def spacing_tight(self: _TFontPosition) -> _TFontPosition:
        """Gets copy of instance with spacing set to tight value"""
        ft = self.copy()
        ft.prop_spacing = CharSpacingKind.TIGHT
        return ft

    @property
    def spacing_normal(self: _TFontPosition) -> _TFontPosition:
        """Gets copy of instance with spacing set to normal value"""
        ft = self.copy()
        ft.prop_spacing = CharSpacingKind.NORMAL
        return ft

    @property
    def spacing_loose(self: _TFontPosition) -> _TFontPosition:
        """Gets copy of instance with spacing set to loose value"""
        ft = self.copy()
        ft.prop_spacing = CharSpacingKind.LOOSE
        return ft

    @property
    def spacing_very_loose(self: _TFontPosition) -> _TFontPosition:
        """Gets copy of instance with spacing set to very loose value"""
        ft = self.copy()
        ft.prop_spacing = CharSpacingKind.VERY_LOOSE
        return ft

    @property
    def pair(self: _TFontPosition) -> _TFontPosition:
        """Gets copy of instance with rotation pair kerning set to ``True``"""
        ft = self.copy()
        ft.prop_pair = True
        return ft

    # endregion Style Properties

    # region Prop Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.CHAR

    @property
    def prop_raise_lower(self) -> int | None:
        """Gets/Sets raise or lower amount, A value of ``0`` means automatic."""
        pv = cast(int, self._get("CharEscapement"))
        if pv is None:
            return None
        # subscript is always a negative value
        return abs(pv)

    @prop_raise_lower.setter
    def prop_raise_lower(self, value: int | None) -> None:
        if value is None:
            self._remove("CharEscapement")
            return

        if value < 0:
            raise ValueError("raise_lower must be a postivie number")
        kind = self.prop_script_kind

        if value == 0:
            # 0 means automatic
            if kind is None:
                self._set("CharEscapement", FontScriptKind.SUPERSCRIPT.value)
                return
            if kind == FontScriptKind.SUBSCRIPT:
                self._set("CharEscapement", FontScriptKind.SUBSCRIPT.value)
                return
            self._set("CharEscapement", FontScriptKind.SUPERSCRIPT.value)
            return

        if kind == FontScriptKind.SUBSCRIPT:
            # subscript is always a negative value
            self._set("CharEscapement", -abs(value))
            return
        self._set("CharEscapement", value)

    @property
    def prop_rel_size(self) -> Intensity | None:
        """Gets/Sets relative font size"""
        pv = cast(int, self._get("CharEscapementHeight"))
        if pv is None:
            return None
        return Intensity(pv)

    @prop_rel_size.setter
    def prop_rel_size(self, value: Intensity | int | None) -> None:
        if value is None:
            self._remove("CharEscapementHeight")
            return
        self._set("CharEscapementHeight", Intensity(int(value)).value)

    @property
    def prop_script_kind(self) -> FontScriptKind | None:
        pv = cast(int, self._get("CharEscapement"))
        if pv is None:
            return None
        if pv == 0:
            return FontScriptKind.NORMAL
        if pv < 0:
            return FontScriptKind.SUBSCRIPT
        return FontScriptKind.SUPERSCRIPT

    @prop_script_kind.setter
    def prop_script_kind(self, value: FontScriptKind | None) -> None:
        if value is None:
            self._remove("CharEscapement")
            return
        height = cast(int, self._get("CharEscapementHeight"))
        if height is None:
            if value == FontScriptKind.NORMAL:
                self._set("CharEscapementHeight", 100)
            else:
                self._set("CharEscapementHeight", FontPosition._DEFAULT_SUPER_SUB_HEIGHT)
        if value == FontScriptKind.NORMAL:
            self._set("CharEscapement", value.value)
            self._set("CharEscapementHeight", 100)
            return
        pv = cast(int, self._get("CharEscapement"))
        if pv is None:
            self._set("CharEscapement", value.value)
            return
        if value == FontScriptKind.SUPERSCRIPT:
            # superscript is always positive value
            self._set("CharEscapement", abs(pv))
        else:
            # subscript is always negative value
            self._set("CharEscapement", -abs(pv))

    @property
    def prop_rotation(self) -> Angle | None:
        """Gets/Sets Font Rotation"""
        pv = cast(int, self._get("CharRotation"))
        if pv is None:
            return None
        return Angle(round(pv / 10))

    @prop_rotation.setter
    def prop_rotation(self, value: int | Angle | None) -> None:
        if value is None:
            self._remove("CharRotation")
            return
        angle = Angle(int(value))

        self._set("CharRotation", angle.value * 10)

    @property
    def prop_scale(self) -> int | None:
        """Gets/Sets scale width"""
        return self._get("CharScaleWidth")

    @prop_scale.setter
    def prop_scale(self, value: int | None) -> None:
        if value is None:
            self._remove("CharScaleWidth")
            return
        if value < 1:
            raise ValueError("Scale cannot be less then 1")

        self._set("CharScaleWidth", value)

    @property
    def prop_fit(self) -> bool | None:
        """Gets/Sets if rotation is fit to line"""
        return self._get("CharRotationIsFitToLine")

    @prop_fit.setter
    def prop_fit(self, value: bool | None) -> None:
        if value is None:
            self._remove("CharRotationIsFitToLine")
            return
        self._set("CharRotationIsFitToLine", value)

    @property
    def prop_spacing(self) -> float | None:
        """This value contains character spacing in point units"""
        pv = self._get("CharKerning")
        if not pv is None:
            if pv == 0.0:
                return 0.0
            return UnitConvert.convert_mm100_pt(pv)
        return None

    @prop_spacing.setter
    def prop_spacing(self, value: float | CharSpacingKind | None) -> None:
        if value is None:
            self._remove("CharKerning")
            return
        self._set("CharKerning", UnitConvert.convert_pt_mm100(float(value)))

    @property
    def prop_pair(self) -> bool | None:
        """Gets/Sets pair kerning"""
        return self._get("CharAutoKerning")

    @prop_pair.setter
    def prop_pair(self, value: bool | None) -> None:
        if value is None:
            self._remove("CharAutoKerning")
            return
        self._set("CharAutoKerning", value)

    # endregion Prop Properties
    @static_prop
    def default() -> FontPosition:  # type: ignore[misc]
        """Gets Font Position default. Static Property."""
        try:
            return FontPosition._DEFAULT_INSTANCE
        except AttributeError:
            fp = FontPosition()
            fp._set("CharEscapement", 0)
            fp._set("CharEscapementHeight", 100)
            fp._set("CharRotation", 0)
            fp._set("CharScaleWidth", 100)
            fp._set("CharRotationIsFitToLine", False)
            fp._set("CharKerning", 0)
            fp._set("CharAutoKerning", True)
            fp._is_default_inst = True
            FontPosition._DEFAULT_INSTANCE = fp
        return FontPosition._DEFAULT_INSTANCE
