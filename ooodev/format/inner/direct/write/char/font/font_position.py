"""
Module for managing character Font position.

.. versionadded:: 0.9.0
"""

# region Import
from __future__ import annotations
from typing import Any, Tuple, Type, cast, overload, TypeVar, TYPE_CHECKING
from enum import Enum

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.loader import lo as mLo
from ooodev.units.angle import Angle
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_pt import UnitPT
from ooodev.utils import props as mProps
from ooodev.utils.data_type.intensity import Intensity

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

# endregion Import

_TFontPosition = TypeVar("_TFontPosition", bound="FontPosition")


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
    Character Font Position.

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together font properties.

    Many properties can be chained together.

    .. seealso::

        - :ref:`help_writer_format_direct_char_font_position`

    .. versionadded:: 0.9.0
    """

    _DEFAULT_SUPER_SUB_HEIGHT = 58

    def __init__(
        self,
        *,
        script_kind: FontScriptKind | None = None,
        raise_lower: int | None = None,
        rel_size: int | Intensity | None = None,
        rotation: int | Angle | None = None,
        scale: int | None = None,
        fit: bool | None = None,
        spacing: CharSpacingKind | float | UnitT | None = None,
        pair: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            script_kind (FontScriptKind, optional): Specifies Superscript/Subscript option.
            raise_lower (int, optional): Specifies raise or Lower as percent value. Set to a value of 0 for automatic.
            rel_size (int, Intensity, optional): Specifies relative Font Size as percent value.
                Set this value to ``0`` for automatic; Otherwise value from ``1`` to ``100``.
            rotation (int, Angle, optional): Specifies the rotation of a character in degrees. Depending on the
                implementation only certain values may be allowed.
            scale (int, optional): Specifies scale width as percent value. Min value is ``1``.
            fit (bool, optional): Specifies if rotation is fit to line.
            spacing (CharSpacingKind, float, UnitT, optional): Specifies character spacing in ``pt`` (point) units
                or :ref:`proto_unit_obj`.
            pair (bool, optional): Specifies pair kerning.

        Returns:
            None:

        See Also:

            - :ref:`help_writer_format_direct_char_font_position`
        """

        super().__init__()
        # superscript and subscript use the same internal properties,CharEscapementHeight, CharEscapement
        if script_kind is not None:
            self.prop_script_kind = script_kind
        if raise_lower is not None:
            self.prop_raise_lower = raise_lower
        if rel_size is not None:
            self.prop_rel_size = rel_size
        if rotation is not None:
            self.prop_rotation = rotation
        if scale is not None:
            self.prop_scale = scale
        if fit is not None:
            self.prop_fit = fit
        if spacing is not None:
            self.prop_spacing = spacing
        if pair is not None:
            self.prop_pair = pair

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.CharacterProperties",
                "com.sun.star.style.CharacterStyle",
                "com.sun.star.style.ParagraphStyle",
            )
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    # region apply()
    @overload
    def apply(self, obj: Any) -> None:  # type: ignore
        ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO object that has supports ``com.sun.star.style.CharacterProperties`` service.

        Returns:
            None:
        """
        super().apply(obj, **kwargs)

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print("FontPosition.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TFontPosition], obj: Any) -> _TFontPosition: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TFontPosition], obj: Any, **kwargs) -> _TFontPosition: ...

    @classmethod
    def from_obj(cls: Type[_TFontPosition], obj: Any, **kwargs) -> _TFontPosition:
        """
        Gets instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.CharacterProperties`` service.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            FontPosition: ``FontPosition`` instance that represents ``obj`` font position.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        def set_prop(key: str, fp: FontPosition):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if val is not None:
                fp._set(key, val)

        set_prop("CharEscapement", inst)
        set_prop("CharAutoEscapement", inst)
        set_prop("CharEscapementHeight", inst)
        set_prop("CharRotation", inst)
        set_prop("CharScaleWidth", inst)
        set_prop("CharRotationIsFitToLine", inst)
        set_prop("CharKerning", inst)
        set_prop("CharAutoKerning", inst)
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()

    def on_property_setting(self, source: Any, event_args: KeyValCancelArgs) -> None:
        """
        Triggers for each property that is set

        Args:
            event_args (KeyValueCancelArgs): Event Args
        """
        # if position is normal then defaults should be set
        if event_args.key == "CharEscapement":
            if self.prop_script_kind == FontScriptKind.NORMAL:
                event_args.value = 0
        elif event_args.key == "CharEscapementHeight":
            if self.prop_script_kind == FontScriptKind.NORMAL:
                event_args.value = 100
        super().on_property_setting(source, event_args)

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
            value (int, optional): relative size value.
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
            value (int, optional): scale width value.
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
            value (bool, optional): Rotation fit value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontPosition: Font with style added or removed
        """
        ft = self.copy()
        ft.prop_fit = value
        return ft

    def fmt_spacing(self: _TFontPosition, value: float | UnitT | None = None) -> _TFontPosition:
        """
        Get copy of instance with spacing set or removed.

        Args:
            value (float, UnitT, optional): The character spacing in ``pt`` (point) units :ref:`proto_unit_obj`.
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
            value (bool, optional): Pair kerning value.
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
    def normal(self: _TFontPosition) -> _TFontPosition:
        """Gets copy of instance set to Position Normal"""
        ft = self.copy()
        ft.prop_script_kind = FontScriptKind.NORMAL
        return ft

    script_kind_normal = normal

    @property
    def superscript(self: _TFontPosition) -> _TFontPosition:
        """Gets copy of instance set to Position Superscript"""
        ft = self.copy()
        ft.prop_script_kind = FontScriptKind.SUPERSCRIPT
        return ft

    script_kind_superscript = superscript

    @property
    def subscript(self: _TFontPosition) -> _TFontPosition:
        """Gets copy of instance set to Position Subscript"""
        ft = self.copy()
        ft.prop_script_kind = FontScriptKind.SUBSCRIPT
        return ft

    script_kind_subscript = subscript

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
    def spacing_very_tight(self: _TFontPosition) -> _TFontPosition:
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
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.CHAR
        return self._format_kind_prop

    @property
    def prop_raise_lower(self) -> int | None:
        """Gets/Sets raise or lower amount, A value of ``0`` means automatic."""
        pv = cast(int, self._get("CharEscapement"))
        return None if pv is None else abs(pv)

    @prop_raise_lower.setter
    def prop_raise_lower(self, value: int | None) -> None:
        if value is None:
            self._remove("CharEscapement")
            self._remove("CharAutoEscapement")
            return

        if value < 0:
            raise ValueError("raise_lower must be a positive number")
        kind = self.prop_script_kind

        if value == 0:
            # 0 means automatic
            self._set("CharAutoEscapement", True)
            if kind is None:
                self._set("CharEscapement", FontScriptKind.SUPERSCRIPT.value)
                return
            if kind == FontScriptKind.SUBSCRIPT:
                self._set("CharEscapement", FontScriptKind.SUBSCRIPT.value)
                return
            self._set("CharEscapement", FontScriptKind.SUPERSCRIPT.value)
            return

        self._set("CharAutoEscapement", False)

        if kind == FontScriptKind.SUBSCRIPT:
            # subscript is always a negative value
            self._set("CharEscapement", -abs(value))
            return
        self._set("CharEscapement", value)

    @property
    def prop_rel_size(self) -> Intensity | None:
        """Gets/Sets relative font size"""
        pv = cast(int, self._get("CharEscapementHeight"))
        return None if pv is None else Intensity(pv)

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
        return FontScriptKind.SUBSCRIPT if pv < 0 else FontScriptKind.SUPERSCRIPT

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
        return None if pv is None else Angle(round(pv / 10))

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
    def prop_spacing(self) -> UnitPT | None:
        """This value contains character spacing in point units"""
        pv = cast(int, self._get("CharKerning"))
        return None if pv is None else UnitPT.from_mm100(pv)

    @prop_spacing.setter
    def prop_spacing(self, value: float | CharSpacingKind | UnitT | None) -> None:
        if value is None:
            self._remove("CharKerning")
            return
        try:
            self._set("CharKerning", value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set("CharKerning", UnitConvert.convert_pt_mm100(float(value)))  # type: ignore

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
    @property
    def default(self: _TFontPosition) -> _TFontPosition:  # type: ignore[misc]
        """Gets Font Position default."""
        # pylint: disable=protected-access
        # pylint: disable=unexpected-keyword-arg
        try:
            return self._default_instance
        except AttributeError:
            fp = self.__class__(_cattribs=self._get_internal_cattribs())  # type: ignore
            fp._set("CharEscapement", 0)
            fp._set("CharEscapementHeight", 100)
            fp._set("CharRotation", 0)
            fp._set("CharScaleWidth", 100)
            fp._set("CharRotationIsFitToLine", False)
            fp._set("CharKerning", 0)
            fp._set("CharAutoKerning", True)
            fp._is_default_inst = True
            self._default_instance = fp
        return self._default_instance
