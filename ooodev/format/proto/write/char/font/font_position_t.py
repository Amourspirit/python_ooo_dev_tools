from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING
import uno


from ooodev.mock.mock_g import DOCS_BUILDING
from ooodev.format.proto.style_t import StyleT

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from ooodev.format.inner.direct.write.char.font.font_position import CharSpacingKind
    from ooodev.format.inner.direct.write.char.font.font_position import FontScriptKind
    from ooodev.units.unit_obj import UnitT
    from ooodev.units.unit_pt import UnitPT
    from ooodev.units.angle import Angle
    from ooodev.utils.data_type.intensity import Intensity
else:
    Protocol = object


class FontPositionT(StyleT, Protocol):
    """Font Position Protocol"""

    def __init__(
        self,
        *,
        script_kind: FontScriptKind | None = ...,
        raise_lower: int | None = ...,
        rel_size: int | Intensity | None = ...,
        rotation: int | Angle | None = ...,
        scale: int | None = ...,
        fit: bool | None = ...,
        spacing: CharSpacingKind | float | UnitT | None = ...,
        pair: bool | None = ...,
    ) -> None:
        """
        Character Font Position.

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
        """
        ...

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls, obj: Any) -> FontPositionT: ...

    @overload
    @classmethod
    def from_obj(cls, obj: Any, **kwargs) -> FontPositionT: ...

    # endregion from_obj()

    # region Format Methods
    def fmt_scrip_kind(self, value: FontScriptKind | None = ...) -> FontPositionT:
        """
        Get copy of instance with superscript/subscript set or removed.

        Args:
            value (FontScriptKind, optional): font script kind.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontPosition: Font with style added or removed
        """
        ...

    def fmt_raise_lower(self, value: int | Intensity | None = ...) -> FontPositionT:
        """
        Get copy of instance with raise/lower set or removed.

        Args:
            value (int, Intensity, optional): Raise or Lower value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontPosition: Font with style added or removed
        """
        ...

    def fmt_rel_size(self, value: int | None = ...) -> FontPositionT:
        """
        Get copy of instance with relative size set or removed.

        Args:
            value (int, optional): relative size value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontPosition: Font with style added or removed
        """
        ...

    def fmt_rotation(self, value: int | Angle | None = ...) -> FontPositionT:
        """
        Get copy of instance with rotation set or removed.

        Args:
            value (int, Angle, optional): The rotation of a character in degrees. Depending on the implementation only certain values may be allowed.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontPosition: Font with style added or removed
        """
        ...

    def fmt_scale(self, value: int | None = ...) -> FontPositionT:
        """
        Get copy of instance with scale width set or removed.

        Args:
            value (int, optional): scale width value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontPosition: Font with style added or removed
        """
        ...

    def fmt_fit(self, value: bool | None = ...) -> FontPositionT:
        """
        Get copy of instance with rotation fit set or removed.

        Args:
            value (bool, optional): Rotation fit value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontPosition: Font with style added or removed
        """
        ...

    def fmt_spacing(self, value: float | UnitT | None = ...) -> FontPositionT:
        """
        Get copy of instance with spacing set or removed.

        Args:
            value (float, UnitT, optional): The character spacing in ``pt`` (point) units :ref:`proto_unit_obj`.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontPosition: Font with style added or removed
        """
        ...

    def fmt_pair(self, value: bool | None = ...) -> FontPositionT:
        """
        Get copy of instance with pair kerning set or removed.

        Args:
            value (bool, optional): Pair kerning value.
                If ``None`` style is removed. Default ``None``

        Returns:
            FontPosition: Font with style added or removed
        """
        ...

    # endregion Format Methods

    # region Style Properties
    @property
    def normal(self) -> FontPositionT:
        """Gets copy of instance set to Position Normal"""
        ...

    script_kind_normal = normal

    @property
    def superscript(self) -> FontPositionT:
        """Gets copy of instance set to Position Superscript"""
        ...

    script_kind_superscript = superscript

    @property
    def subscript(self) -> FontPositionT:
        """Gets copy of instance set to Position Subscript"""
        ...

    script_kind_subscript = subscript

    @property
    def raise_lower_auto(self) -> FontPositionT:
        """Gets copy of instance with raise/lower set to automatic"""
        ...

    @property
    def rotation_none(self) -> FontPositionT:
        """Gets copy of instance with rotation set to ``0``"""
        ...

    @property
    def rotation_90(self) -> FontPositionT:
        """Gets copy of instance with rotation set to ``90``"""
        ...

    @property
    def rotation_270(self) -> FontPositionT:
        """Gets copy of instance with rotation set to ``270``"""
        ...

    @property
    def fit(self) -> FontPositionT:
        """Gets copy of instance with rotation fit to line set to ``True``"""
        ...

    @property
    def spacing_very_tight(self) -> FontPositionT:
        """Gets copy of instance with spacing set to very tight value"""
        ...

    @property
    def spacing_tight(self) -> FontPositionT:
        """Gets copy of instance with spacing set to tight value"""
        ...

    @property
    def spacing_normal(self) -> FontPositionT:
        """Gets copy of instance with spacing set to normal value"""
        ft = self.copy()
        ft.prop_spacing = CharSpacingKind.NORMAL
        return ft

    @property
    def spacing_loose(self) -> FontPositionT:
        """Gets copy of instance with spacing set to loose value"""
        ...

    @property
    def spacing_very_loose(self) -> FontPositionT:
        """Gets copy of instance with spacing set to very loose value"""
        ...

    @property
    def pair(self) -> FontPositionT:
        """Gets copy of instance with rotation pair kerning set to ``True``"""
        ...

    # endregion Style Properties

    # region Prop Properties

    @property
    def prop_raise_lower(self) -> int | None:
        """Gets/Sets raise or lower amount, A value of ``0`` means automatic."""
        ...

    @prop_raise_lower.setter
    def prop_raise_lower(self, value: int | None) -> None: ...

    @property
    def prop_rel_size(self) -> Intensity | None:
        """Gets/Sets relative font size"""
        ...

    @prop_rel_size.setter
    def prop_rel_size(self, value: Intensity | int | None) -> None: ...

    @property
    def prop_script_kind(self) -> FontScriptKind | None: ...

    @prop_script_kind.setter
    def prop_script_kind(self, value: FontScriptKind | None) -> None: ...

    @property
    def prop_rotation(self) -> Angle | None:
        """Gets/Sets Font Rotation"""
        ...

    @prop_rotation.setter
    def prop_rotation(self, value: int | Angle | None) -> None: ...

    @property
    def prop_scale(self) -> int | None:
        """Gets/Sets scale width"""
        ...

    @prop_scale.setter
    def prop_scale(self, value: int | None) -> None: ...

    @property
    def prop_fit(self) -> bool | None:
        """Gets/Sets if rotation is fit to line"""
        ...

    @prop_fit.setter
    def prop_fit(self, value: bool | None) -> None: ...

    @property
    def prop_spacing(self) -> UnitPT | None:
        """This value contains character spacing in point units"""
        ...

    @prop_spacing.setter
    def prop_spacing(self, value: float | CharSpacingKind | UnitT | None) -> None: ...

    @property
    def prop_pair(self) -> bool | None:
        """Gets/Sets pair kerning"""
        ...

    @prop_pair.setter
    def prop_pair(self, value: bool | None) -> None: ...

    # endregion Prop Properties
    @property
    def default(self) -> FontPositionT:  # type: ignore[misc]
        """Gets Font Position default."""
        ...
