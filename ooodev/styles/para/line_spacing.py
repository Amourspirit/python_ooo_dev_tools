"""
Modele for managing paragraph Spacing.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, cast, overload
from enum import Enum
from numbers import Real

from ...exceptions import ex as mEx
from ...meta.static_prop import static_prop
from ...utils import lo as mLo
from ...utils import props as mProps
from ..style_base import StyleMulti
from ..structs import line_spacing as mLs
from ..kind.style_kind import StyleKind


class ModeKind(Enum):
    """Mode Kinde"""

    SINGLE = (0, 0)
    """Single"""
    PORPORTINAL = (1, 0)
    """Portortinal"""
    LINES_1_15 = (2, 0)
    """1.15 Lines"""
    LINES_15 = (3, 0)
    """1.5 Lines"""
    DOUBLE = (4, 0)
    """Double"""
    AT_LEAST = (5, 1)
    """At Least"""
    LEADING = (6, 2)
    """Leading"""
    FIXED = (7, 3)
    """Fixed"""

    def __int__(self) -> int:
        return self.value[1]


class LineSpacing(StyleMulti):
    """
    Paragraph Line Spacing

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``style_`` can be used to chain together Padding properties.

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

    # region init

    def __init__(
        self,
        mode: ModeKind | None = None,
        value: Real | None = None,
        active_ln_spacing: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            mode (ModeKind, optional): Determines the left margin of the paragraph (in mm units).
            value (Real, optional): Value of line spacing. Only applies when ``ModeKind`` is ``PORPORTINAL``, ``AT_LEAST``, ``LEADING``, or ``FIXED``.
            active_ln_spacing (bool, optional): Determines active page line-spacing.
        Returns:
            None:

        Note:
            When ``mode`` is ``ModeKind.AT_LEAST``, ``ModeKind.LEADING``, or ``ModeKind.FIXED``
            then the units are mm units (as float).

            When ``mode`` is ``ModeKind.PORPORTINAL`` then the unit is percentage (as int).
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        init_vals = {}

        ls = None
        if not mode is None:
            if mode == ModeKind.SINGLE:
                ls = mLs.LineSpacing(int(mode), 100)
            elif mode == ModeKind.LINES_1_15:
                ls = mLs.LineSpacing(int(mode), 115)
            elif mode == ModeKind.LINES_15:
                ls = mLs.LineSpacing(int(mode), 150)
            elif mode == ModeKind.DOUBLE:
                ls = mLs.LineSpacing(int(mode), 200)
            elif mode == ModeKind.PORPORTINAL:
                # value is considered a percentage
                if not isinstance(value, int):
                    raise TypeError("Value must be a integer when mode is PORPORTINAL")
                ls = mLs.LineSpacing(int(mode), value)
            elif mode in (ModeKind.AT_LEAST, ModeKind.LEADING, ModeKind.FIXED):
                if not isinstance(value, Real):
                    raise TypeError("Value must be a integer or float when mode is AT_LEAST or LEADING or FIXED")
                # value is considered to be mm units
                ls = mLs.LineSpacing(int(mode), round(value * 100))

        if not active_ln_spacing is None:
            init_vals["ParaRegisterModeActive"] = active_ln_spacing

        super().__init__(**init_vals)
        if not mode is None:
            self._set_style("line_spacing", ls, "ParaLineSpacing", keys={"spacing": "ParaLineSpacing"})

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.style.ParagraphProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.ParagraphProperties",)

    @overload
    def apply_style(self, obj: object) -> None:
        ...

    def apply_style(self, obj: object, **kwargs) -> None:
        """
        Applies writing mode to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            None:
        """
        try:
            super().apply_style(obj)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__}.apply_style(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    @staticmethod
    def from_obj(obj: object) -> LineSpacing:
        """
        Gets instance from object

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Raises:
            NotSupportedServiceError: If ``obj`` does not support  ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            LineSpacing: ``LineSpacing`` instance that represents ``obj`` line spacing.
        """
        inst = LineSpacing()
        if not inst._is_valid_service(obj):
            raise mEx.NotSupportedServiceError(inst._supported_services()[0])

        def set_prop(key: str, indent: LineSpacing):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if not val is None:
                indent._set(key, val)

        set_prop("ParaRegisterModeActive", inst)

        ls = mProps.Props.get(obj, "ParaLineSpacing", None)
        if not ls is None:
            inst._set_style("line_spacing", ls, keys={"spacing": "ParaLineSpacing"})
        return inst

    # endregion methods

    # region properties
    @property
    def prop_style_kind(self) -> StyleKind:
        """Gets the kind of style"""
        return StyleKind.PARA

    @property
    def prop_active_ln_spacing(self) -> bool | None:
        """Gets/Sets active page line-spacing."""
        return self._get("ParaRegisterModeActive")

    @prop_active_ln_spacing.setter
    def prop_active_ln_spacing(self, value: bool | None):
        if value is None:
            self._remove("ParaRegisterModeActive")
            return
        self._set("ParaRegisterModeActive", value)

    @static_prop
    def default(cls) -> LineSpacing:
        """Gets ``LineSpacing`` default. Static Property."""
        if cls._DEFAULT is None:
            cls._DEFAULT = LineSpacing(mode=ModeKind.SINGLE)
        return cls._DEFAULT

    # endregion properties
