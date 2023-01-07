"""
Module for managing character padding.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import cast

from ...exceptions import ex as mEx
from ...meta.static_prop import static_prop
from ...utils import info as mInfo
from ...utils import lo as mLo
from ...utils import props as mProps
from ..style_base import StyleBase
from ..kind.style_kind import StyleKind


class Padding(StyleBase):
    """
    Character Padding

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``style_`` can be used to chain together Sides properties.

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    # this class also set Borders Padding in borders.Border class.

    def __init__(
        self,
        left: float | None = None,
        right: float | None = None,
        top: float | None = None,
        bottom: float | None = None,
        padding_all: float | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (float, optional): Character left padding (in mm units).
            right (float, optional): Character right padding (in mm units).
            top (float, optional): Character top padding (in mm units).
            bottom (float, optional): Character bottom padding (in mm units).
            padding_all (float, optional): Character left, right, top, bottom padding (in mm units). If argument is present then ``left``, ``right``, ``top``, and ``bottom`` arguments are ignored.

        Raises:
            ValueError: If any argument value is less than zero.

        Returns:
            None:
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        init_vals = {}

        def validate(val: float | None) -> None:
            if val is not None:
                if val < 0.0:
                    raise ValueError("padding values must be positive values")

        def set_val(key, value) -> None:
            nonlocal init_vals
            if not value is None:
                init_vals[key] = round(value * 100)

        validate(left)
        validate(right)
        validate(top)
        validate(bottom)
        validate(padding_all)
        char_attrs = (
            "CharLeftBorderDistance",
            "CharRightBorderDistance",
            "CharTopBorderDistance",
            "CharBottomBorderDistance",
        )
        if padding_all is None:
            for key, value in zip(char_attrs, (left, right, top, bottom)):
                set_val(key, value)
        else:
            for key in char_attrs:
                set_val(key, padding_all)

        super().__init__(**init_vals)

    def _is_supported(self, obj: object) -> bool:
        return mInfo.Info.support_service(obj, "com.sun.star.style.CharacterProperties")

    def apply_style(self, obj: object, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.CharacterProperties`` service.
            kwargs (Any, optional): Expandable list of key value pairs that may be used in child classes.

        Returns:
            None:
        """
        if self._is_supported(obj):
            try:
                super().apply_style(obj)
            except mEx.MultiError as e:
                mLo.Lo.print(f"{self.__class__}.apply_style(): Unable to set Property")
                for err in e.errors:
                    mLo.Lo.print(f"  {err}")
        else:
            mLo.Lo.print(f'{self.__class__}.apply_style(): "com.sun.star.style.CharacterProperties" not supported')
        return None

    @staticmethod
    def from_obj(obj: object) -> Padding:
        """
        Gets Padding instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.CharacterProperties`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support  ``com.sun.star.style.CharacterProperties`` service.

        Returns:
            Padding: Padding that represents ``obj`` padding.
        """
        if not mInfo.Info.support_service(obj, "com.sun.star.style.CharacterProperties"):
            raise mEx.NotSupportedServiceError("com.sun.star.style.CharacterProperties")
        pd = Padding()
        pd._set("CharLeftBorderDistance", int(mProps.Props.get(obj, "CharLeftBorderDistance")))
        pd._set("CharRightBorderDistance", int(mProps.Props.get(obj, "CharRightBorderDistance")))
        pd._set("CharTopBorderDistance", int(mProps.Props.get(obj, "CharTopBorderDistance")))
        pd._set("CharBottomBorderDistance", int(mProps.Props.get(obj, "CharBottomBorderDistance")))
        return pd

    # region Syle methods
    def style_padding_all(self, value: float | None) -> Padding:
        """
        Gets copy of instance with left, right, top, bottom sides set or removed

        Args:
            value (float | None): padding value

        Returns:
            Padding: Padding instance
        """
        cp = self.copy()
        cp.prop_top = value
        cp.prop_bottom = value
        cp.prop_left = value
        cp.prop_right = value
        return cp

    def style_top(self, value: float | None) -> Padding:
        """
        Gets a copy of instance with top side set or removed

        Args:
            value (float | None): padding value

        Returns:
            Padding: Padding instance
        """
        cp = self.copy()
        cp.prop_top = value
        return cp

    def style_bottom(self, value: float | None) -> Padding:
        """
        Gets a copy of instance with bottom side set or removed

        Args:
            value (float | None): padding value

        Returns:
            Padding: Padding instance
        """
        cp = self.copy()
        cp.prop_bottom = value
        return cp

    def style_left(self, value: float | None) -> Padding:
        """
        Gets a copy of instance with left side set or removed

        Args:
            value (float | None): padding value

        Returns:
            Padding: Padding instance
        """
        cp = self.copy()
        cp.prop_left = value
        return cp

    def style_right(self, value: float | None) -> Padding:
        """
        Gets a copy of instance with right side set or removed

        Args:
            value (float | None): padding value

        Returns:
            Padding: Padding instance
        """
        cp = self.copy()
        cp.prop_right = value
        return cp

    # endregion style methods

    @property
    def prop_style_kind(self) -> StyleKind:
        """Gets the kind of style"""
        return StyleKind.CHAR

    @property
    def prop_left(self) -> float | None:
        """Gets/Sets paragraph left padding (in mm units)."""
        pv = cast(int, self._get("CharLeftBorderDistance"))
        if pv is None:
            return None
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_left.setter
    def prop_left(self, value: float | None):
        if value is None:
            self._remove("CharLeftBorderDistance")
            return
        self._set("CharLeftBorderDistance", round(value * 100))

    @property
    def prop_right(self) -> float | None:
        """Gets/Sets paragraph right padding (in mm units)."""
        pv = cast(int, self._get("CharRightBorderDistance"))
        if pv is None:
            return None
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_right.setter
    def prop_right(self, value: float | None):
        if value is None:
            self._remove("CharRightBorderDistance")
            return
        self._set("CharRightBorderDistance", round(value * 100))

    @property
    def prop_top(self) -> float | None:
        """Gets/Sets paragraph top padding (in mm units)."""
        pv = cast(int, self._get("CharTopBorderDistance"))
        if pv is None:
            return None
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_top.setter
    def prop_top(self, value: float | None):
        if value is None:
            self._remove("CharTopBorderDistance")
            return
        self._set("CharTopBorderDistance", round(value * 100))

    @property
    def prop_bottom(self) -> float | None:
        """Gets/Sets paragraph bottom padding (in mm units)."""
        pv = cast(int, self._get("CharBottomBorderDistance"))
        if pv is None:
            return None
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_bottom.setter
    def prop_bottom(self, value: float | None):
        if value is None:
            self._remove("CharBottomBorderDistance")
            return
        self._set("CharBottomBorderDistance", round(value * 100))

    @static_prop
    def default(cls) -> Padding:
        """Gets Padding default. Static Property."""
        if cls._DEFAULT is None:
            cls._DEFAULT = Padding(padding_all=0.0)
        return cls._DEFAULT
