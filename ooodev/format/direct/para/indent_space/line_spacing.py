"""
Modele for managing paragraph Line Spacing.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, cast, overload
from numbers import Real

from .....events.args.cancel_event_args import CancelEventArgs
from .....exceptions import ex as mEx
from .....meta.static_prop import static_prop
from .....utils import lo as mLo
from .....utils import props as mProps
from ....style_base import StyleMulti
from ...structs import line_spacing_struct as mLs
from ...structs.line_spacing_struct import ModeKind as ModeKind
from ....kind.format_kind import FormatKind


class ParaLineSpace(mLs.LineSpacingStruct):
    """Represents a Line spacing Struct for use with paragraphs"""

    def _get_property_name(self) -> str:
        return "ParaLineSpacing"

    def _supported_services(self) -> Tuple[str, ...]:
        # will affect apply() on parent class.
        return ("com.sun.star.style.ParagraphProperties",)

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(event)

    @static_prop
    def default() -> ParaLineSpace:  # type: ignore[misc]
        """Gets empty Line Spacing. Static Property."""
        try:
            return ParaLineSpace._DEFAULT_INST
        except AttributeError:
            ParaLineSpace._DEFAULT_INST = ParaLineSpace(mLs.ModeKind.SINGLE, 0)
            ParaLineSpace._DEFAULT_INST._is_default_inst = True
        return ParaLineSpace._DEFAULT_INST


class LineSpacing(StyleMulti):
    """
    Paragraph Line Spacing

    Any properties starting with ``prop_`` set or get current instance values.
    """

    # region init

    def __init__(
        self,
        mode: ModeKind | None = None,
        value: Real = 0,
        active_ln_spacing: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            mode (ModeKind, optional): Determines the mode that is use to apply units.
            value (Real, optional): Value of line spacing. Only applies when ``ModeKind`` is ``PROPORTIONAL``, ``AT_LEAST``, ``LEADING``, or ``FIXED``.
            active_ln_spacing (bool, optional): Determines active page line-spacing.
        Returns:
            None:

        Note:
            If ``ModeKind`` is ``SINGLE``, ``LINE_1_15``, ``LINE_1_5``, or ``DOUBLE`` then ``value`` is ignored.

            If ``ModeKind`` is ``AT_LEAST``, ``LEADING``, or ``FIXED`` then ``value`` is a float (``in mm uints``).

            If ``ModeKind`` is ``PROPORTIONAL`` then value is an int representing percentage.
            For example ``95`` equals ``95%``, ``130`` equals ``130%``
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        super().__init__()

        ls = None
        if not mode is None:
            ls = ParaLineSpace(mode=mode, value=value)

        self._mode = mode
        self._value = value

        if not active_ln_spacing is None:
            self._set("ParaRegisterModeActive", active_ln_spacing)

        if not ls is None:
            self._set_style("line_spacing", ls, *ls.get_attrs())

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        """
        Gets a tuple of supported services (``com.sun.star.style.ParagraphProperties``,)

        Returns:
            Tuple[str, ...]: Supported services
        """
        return ("com.sun.star.style.ParagraphProperties",)

    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(event)

    # region apply()
    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies writing mode to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            None:
        """
        try:
            super().apply(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()

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
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedServiceError(inst._supported_services()[0])

        def set_prop(key: str, ls_inst: LineSpacing):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if not val is None:
                ls_inst._set(key, val)

        set_prop("ParaRegisterModeActive", inst)

        ls = mProps.Props.get(obj, "ParaLineSpacing", None)
        if not ls is None:
            inst._set_style("line_spacing", ls, *ls.get_attrs())
        return inst

    # endregion methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA

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

    @property
    def prop_mode(self) -> ModeKind | None:
        """Gets the mode that is use to apply units."""
        info = self._get_style("line_spacing")
        if info is None:
            return None
        ls = cast(ParaLineSpace, info[0])
        return ls.prop_mode

    @property
    def prop_value(self) -> Real | None:
        """Gets the Value of line spacing."""
        info = self._get_style("line_spacing")
        if info is None:
            return None
        ls = cast(ParaLineSpace, info[0])
        return ls.prop_value

    @static_prop
    def default() -> LineSpacing:  # type: ignore[misc]
        """Gets ``LineSpacing`` default. Static Property."""
        try:
            return LineSpacing._DEFAULT_INST
        except AttributeError:
            LineSpacing._DEFAULT_INST = LineSpacing(mode=ModeKind.SINGLE)
            LineSpacing._DEFAULT_INST._is_default_inst = True
        return LineSpacing._DEFAULT_INST

    # endregion properties
