"""
Modele for managing paragraph Line Spacing.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Any, Tuple, cast, overload, Type, TypeVar
from numbers import Real

from .....events.args.cancel_event_args import CancelEventArgs
from .....exceptions import ex as mEx
from .....utils import lo as mLo
from .....utils import props as mProps
from ....style_base import StyleMulti
from ...structs import line_spacing_struct as mLs
from ...structs.line_spacing_struct import ModeKind as ModeKind
from ....kind.format_kind import FormatKind

from ooo.dyn.style.line_spacing import LineSpacing as UnoLineSpacing

_TLineSpacing = TypeVar(name="_TLineSpacing", bound="LineSpacing")


class LineSpacing(StyleMulti):
    """
    Paragraph Line Spacing

    Any properties starting with ``prop_`` set or get current instance values.
    """

    # region init

    def __init__(
        self,
        *,
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
            ls = mLs.LineSpacingStruct(
                mode=mode,
                value=value,
                _cattribs=self._get_ls_cattribs(),
            )

        if not active_ln_spacing is None:
            self._set("ParaRegisterModeActive", active_ln_spacing)

        if not ls is None:
            self._set_style("line_spacing", ls, *ls.get_attrs())

    # endregion init

    # region Internal Methods
    def _get_ls_cattribs(self) -> dict:
        return {
            "_supported_services_values": self._supported_services(),
            "_property_name": "ParaLineSpacing",
        }

    # endregion Internal Methods

    # region Overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.ParagraphProperties",
                "com.sun.star.style.ParagraphStyle",
            )
        return self._supported_services_values

    def _on_modifing(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(source, event)

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

    # endregion Overrides

    # region static methods

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TLineSpacing], obj: object) -> _TLineSpacing:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TLineSpacing], obj: object, **kwargs) -> _TLineSpacing:
        ...

    @classmethod
    def from_obj(cls: Type[_TLineSpacing], obj: object, **kwargs) -> _TLineSpacing:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            LineSpacing: ``LineSpacing`` instance that represents ``obj`` line spacing.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        def set_prop(key: str, ls_inst: LineSpacing):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if not val is None:
                ls_inst._set(key, val)

        set_prop("ParaRegisterModeActive", inst)

        pls = cast(UnoLineSpacing, mProps.Props.get(obj, "ParaLineSpacing", None))
        ls = mLs.LineSpacingStruct.from_uno_struct(
            pls,
            _cattribs={"_supported_services_values": inst._supported_services(), "_property_name": "ParaLineSpacing"},
        )
        if not ls is None:
            inst._set_style("line_spacing", ls, *ls.get_attrs())
        return inst

    # endregion from_obj()

    # endregion static methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA
        return self._format_kind_prop

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
        return self.prop_inner.prop_mode

    @property
    def prop_value(self) -> Real | None:
        """Gets the Value of line spacing."""
        return self.prop_inner.prop_value

    @property
    def prop_inner(self) -> mLs.LineSpacingStruct:
        """Gets Line Spacing instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(mLs.LineSpacingStruct, self._get_style_inst("line_spacing"))
        return self._direct_inner

    @property
    def default(self: _TLineSpacing) -> _TLineSpacing:
        """Gets ``LineSpacing`` default."""
        try:
            return self._default_inst
        except AttributeError:
            self._default_inst = self.__class__(mode=ModeKind.SINGLE, _cattribs=self._get_internal_cattribs())
            self._default_inst._is_default_inst = True
        return self._default_inst

    # endregion properties
