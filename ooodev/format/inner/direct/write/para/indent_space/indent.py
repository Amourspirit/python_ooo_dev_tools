"""
Module for managing paragraph padding.

.. versionadded:: 0.9.0
"""
# region Import
from __future__ import annotations
from typing import Any, Tuple, cast, overload, Type, TypeVar

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.proto.unit_obj import UnitObj
from ooodev.utils import lo as mLo
from ooodev.utils import props as mProps
from ooodev.units.unit_mm import UnitMM
from ooodev.units.unit_convert import UnitConvert
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase

# endregion Import

_TIndent = TypeVar(name="_TIndent", bound="Indent")


class Indent(StyleBase):
    """
    Paragraph Indent

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        *,
        before: float | UnitObj | None = None,
        after: float | UnitObj | None = None,
        first: float | UnitObj | None = None,
        auto: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            before (float, UnitObj, optional): Determines the left margin of the paragraph (in ``mm`` units)
                or :ref:`proto_unit_obj`.
            after (float, UnitObj, optional): Determines the right margin of the paragraph (in ``mm`` units)
                or :ref:`proto_unit_obj`.
            first (float, UnitObj, optional): specifies the indent for the first line (in ``mm`` units)
                or :ref:`proto_unit_obj`.
            auto (bool, optional): Determines if the first line should be indented automatically.

        Returns:
            None:
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        super().__init__()

        if not before is None:
            self.prop_before = before
        if not after is None:
            self.prop_after = after
        if not first is None:
            self.prop_first = first
        if not auto is None:
            self.prop_auto = auto

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.ParagraphProperties",
                "com.sun.star.style.ParagraphStyle",
            )
        return self._supported_services_values

    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

    # region apply()
    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies writing mode to ``obj``

        Args:
            obj (object): UNO object.

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

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TIndent], obj: object) -> _TIndent:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TIndent], obj: object, **kwargs) -> _TIndent:
        ...

    @classmethod
    def from_obj(cls: Type[_TIndent], obj: object, **kwargs) -> _TIndent:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Indent: ``Indent`` instance that represents ``obj`` writing mode.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        def set_prop(key: str, indent: Indent):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if not val is None:
                indent._set(key, val)

        set_prop("ParaLeftMargin", inst)
        set_prop("ParaRightMargin", inst)
        set_prop("ParaFirstLineIndent", inst)
        set_prop("ParaIsAutoFirstLineIndent", inst)
        return inst

    # endregion from_obj()

    # endregion methods

    # region style methods
    def fmt_before(self: _TIndent, value: float | UnitObj | None) -> _TIndent:
        """
        Gets a copy of instance with before margin set or removed

        Args:
            value (float, UnitObj, optional): Margin value (in ``mm`` units) or :ref:`proto_unit_obj`.

        Returns:
            Indent: Indent instance
        """
        cp = self.copy()
        cp.prop_before = value
        return cp

    def fmt_after(self: _TIndent, value: float | UnitObj | None) -> _TIndent:
        """
        Gets a copy of instance with after margin set or removed

        Args:
            value (float, UnitObj, optional): Margin value (in mm units) or :ref:`proto_unit_obj`.

        Returns:
            Indent: Indent instance
        """
        cp = self.copy()
        cp.prop_after = value
        return cp

    def fmt_first(self: _TIndent, value: float | UnitObj | None) -> _TIndent:
        """
        Gets a copy of instance with first indent margin set or removed

        Args:
            value (float, UnitObj, optional): Margin value (in mm units) or :ref:`proto_unit_obj`.

        Returns:
            Indent: Indent instance
        """
        cp = self.copy()
        cp.prop_after = value
        return cp

    def fmt_auto(self: _TIndent, value: bool | None) -> _TIndent:
        """
        Gets a copy of instance with auto set or removed

        Args:
            value (bool | None): Auto value.

        Returns:
            Indent: Indent instance
        """
        cp = self.copy()
        cp.prop_auto = value
        return cp

    # endregion style methods

    # region Style Properties
    @property
    def auto(self: _TIndent) -> _TIndent:
        """Gets copy of instance with auto set"""
        cp = self.copy()
        cp.prop_auto = True
        return cp

    # endregion Style Properties

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
    def prop_before(self) -> UnitMM | None:
        """Gets/Sets the left margin of the paragraph (in mm units)."""
        pv = cast(int, self._get("ParaLeftMargin"))
        if pv is None:
            return None
        return UnitMM.from_mm100(pv)

    @prop_before.setter
    def prop_before(self, value: float | UnitObj | None):
        if value is None:
            self._remove("ParaLeftMargin")
            return
        try:
            self._set("ParaLeftMargin", value.get_value_mm100())
        except AttributeError:
            self._set("ParaLeftMargin", UnitConvert.convert_mm_mm100(value))

    @property
    def prop_after(self) -> UnitMM | None:
        """Gets/Sets the right margin of the paragraph (in mm units)."""
        pv = cast(int, self._get("ParaRightMargin"))
        if pv is None:
            return None
        return UnitMM.from_mm100(pv)

    @prop_after.setter
    def prop_after(self, value: float | UnitObj | None):
        if value is None:
            self._remove("ParaRightMargin")
            return
        try:
            self._set("ParaRightMargin", value.get_value_mm100())
        except AttributeError:
            self._set("ParaRightMargin", UnitConvert.convert_mm_mm100(value))

    @property
    def prop_first(self) -> UnitMM | None:
        """Gets/Sets the indent for the first line (in mm units)."""
        pv = cast(int, self._get("ParaFirstLineIndent"))
        if pv is None:
            return None
        return UnitMM.from_mm100(pv)

    @prop_first.setter
    def prop_first(self, value: float | UnitObj | None):
        if value is None:
            self._remove("ParaFirstLineIndent")
            return
        try:
            self._set("ParaFirstLineIndent", value.get_value_mm100())
        except AttributeError:
            self._set("ParaFirstLineIndent", UnitConvert.convert_mm_mm100(value))

    @property
    def prop_auto(self) -> bool | None:
        """Gets/Sets if the first line should be indented automatically"""
        return self._get("ParaIsAutoFirstLineIndent")

    @prop_auto.setter
    def prop_auto(self, value: bool | None):
        if value is None:
            self._remove("ParaIsAutoFirstLineIndent")
            return
        self._set("ParaIsAutoFirstLineIndent", value)

    @property
    def default(self: _TIndent) -> _TIndent:
        """Gets ``Indent`` default."""
        try:
            return self._default_inst
        except AttributeError:
            self._default_inst = self.__class__(
                before=0.0, after=0.0, first=0.0, auto=False, _cattribs=self._get_internal_cattribs()
            )
            self._default_inst._is_default_inst = True
        return self._default_inst

    # endregion properties
