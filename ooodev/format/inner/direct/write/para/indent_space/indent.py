"""
Module for managing paragraph padding.

.. versionadded:: 0.9.0
"""

# region Import
from __future__ import annotations
from typing import Any, Tuple, cast, overload, Type, TypeVar, TYPE_CHECKING

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.loader import lo as mLo
from ooodev.units.unit_convert import UnitConvert
from ooodev.units.unit_mm import UnitMM
from ooodev.utils import props as mProps

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

# endregion Import

_TIndent = TypeVar("_TIndent", bound="Indent")


class Indent(StyleBase):
    """
    Paragraph Indent

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. seealso::

        - :ref:`help_writer_format_direct_para_indent_spacing`

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        *,
        before: float | UnitT | None = None,
        after: float | UnitT | None = None,
        first: float | UnitT | None = None,
        auto: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            before (float, UnitT, optional): Determines the left margin of the paragraph (in ``mm`` units)
                or :ref:`proto_unit_obj`.
            after (float, UnitT, optional): Determines the right margin of the paragraph (in ``mm`` units)
                or :ref:`proto_unit_obj`.
            first (float, UnitT, optional): specifies the indent for the first line (in ``mm`` units)
                or :ref:`proto_unit_obj`.
            auto (bool, optional): Determines if the first line should be indented automatically.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_direct_para_indent_spacing`
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        super().__init__()

        if before is not None:
            self.prop_before = before
        if after is not None:
            self.prop_after = after
        if first is not None:
            self.prop_first = first
        if auto is not None:
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
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, **kwargs) -> None: ...

    def apply(self, obj: Any, **kwargs) -> None:
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
    def from_obj(cls: Type[_TIndent], obj: Any) -> _TIndent: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TIndent], obj: Any, **kwargs) -> _TIndent: ...

    @classmethod
    def from_obj(cls: Type[_TIndent], obj: Any, **kwargs) -> _TIndent:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Indent: ``Indent`` instance that represents ``obj`` writing mode.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        def set_prop(key: str, indent: Indent):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if val is not None:
                indent._set(key, val)

        set_prop("ParaLeftMargin", inst)
        set_prop("ParaRightMargin", inst)
        set_prop("ParaFirstLineIndent", inst)
        set_prop("ParaIsAutoFirstLineIndent", inst)
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()

    # endregion methods

    # region style methods
    def fmt_before(self: _TIndent, value: float | UnitT | None) -> _TIndent:
        """
        Gets a copy of instance with before margin set or removed

        Args:
            value (float, UnitT, optional): Margin value (in ``mm`` units) or :ref:`proto_unit_obj`.

        Returns:
            Indent: Indent instance
        """
        cp = self.copy()
        cp.prop_before = value
        return cp

    def fmt_after(self: _TIndent, value: float | UnitT | None) -> _TIndent:
        """
        Gets a copy of instance with after margin set or removed

        Args:
            value (float, UnitT, optional): Margin value (in mm units) or :ref:`proto_unit_obj`.

        Returns:
            Indent: Indent instance
        """
        cp = self.copy()
        cp.prop_after = value
        return cp

    def fmt_first(self: _TIndent, value: float | UnitT | None) -> _TIndent:
        """
        Gets a copy of instance with first indent margin set or removed

        Args:
            value (float, UnitT, optional): Margin value (in mm units) or :ref:`proto_unit_obj`.

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
        return None if pv is None else UnitMM.from_mm100(pv)

    @prop_before.setter
    def prop_before(self, value: float | UnitT | None):
        if value is None:
            self._remove("ParaLeftMargin")
            return
        try:
            self._set("ParaLeftMargin", value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set("ParaLeftMargin", UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def prop_after(self) -> UnitMM | None:
        """Gets/Sets the right margin of the paragraph (in mm units)."""
        pv = cast(int, self._get("ParaRightMargin"))
        return None if pv is None else UnitMM.from_mm100(pv)

    @prop_after.setter
    def prop_after(self, value: float | UnitT | None):
        if value is None:
            self._remove("ParaRightMargin")
            return
        try:
            self._set("ParaRightMargin", value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set("ParaRightMargin", UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def prop_first(self) -> UnitMM | None:
        """Gets/Sets the indent for the first line (in mm units)."""
        pv = cast(int, self._get("ParaFirstLineIndent"))
        return None if pv is None else UnitMM.from_mm100(pv)

    @prop_first.setter
    def prop_first(self, value: float | UnitT | None):
        if value is None:
            self._remove("ParaFirstLineIndent")
            return
        try:
            self._set("ParaFirstLineIndent", value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set("ParaFirstLineIndent", UnitConvert.convert_mm_mm100(value))  # type: ignore

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
                before=0.0, after=0.0, first=0.0, auto=False, _cattribs=self._get_internal_cattribs()  # type: ignore
            )
            self._default_inst._is_default_inst = True
        return self._default_inst

    # endregion properties
