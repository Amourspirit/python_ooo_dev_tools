"""
Module for managing paragraph spacing.

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


_TSpacing = TypeVar(name="_TSpacing", bound="Spacing")


class Spacing(StyleBase):
    """
    Paragraph Spacing

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
        above: float | UnitT | None = None,
        below: float | UnitT | None = None,
        style_no_space: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            above (float, UnitT, optional): Determines the top margin of the paragraph (in ``mm`` units)
                or :ref:`proto_unit_obj`.
            below (float, UnitT, optional): Determines the bottom margin of the paragraph (in ``mm`` units)
                or :ref:`proto_unit_obj`.
            style_no_space (bool, optional): Do not add space between paragraphs of the same style.

        Returns:
            None:

        See Also:

            - :ref:`help_writer_format_direct_para_indent_spacing`
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        super().__init__()

        if above is not None:
            self.prop_above = above
        if below is not None:
            self.prop_below = below
        if style_no_space is not None:
            self.prop_style_no_space = style_no_space

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
    def from_obj(cls: Type[_TSpacing], obj: Any) -> _TSpacing: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TSpacing], obj: Any, **kwargs) -> _TSpacing: ...

    @classmethod
    def from_obj(cls: Type[_TSpacing], obj: Any, **kwargs) -> _TSpacing:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Spacing: ``Spacing`` instance that represents ``obj`` spacing.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        def set_prop(key: str, indent: Spacing):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if val is not None:
                indent._set(key, val)

        set_prop("ParaTopMargin", inst)
        set_prop("ParaBottomMargin", inst)
        set_prop("ParaContextMargin", inst)
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()

    # endregion methods

    # region style methods
    def fmt_above(self: _TSpacing, value: float | UnitT | None) -> _TSpacing:
        """
        Gets a copy of instance with above margin set or removed

        Args:
            value (float, UnitT, optional): Margin value (in ``mm`` units) or :ref:`proto_unit_obj`.

        Returns:
            Spacing: Indent instance
        """
        cp = self.copy()
        cp.prop_above = value
        return cp

    def fmt_below(self: _TSpacing, value: float | UnitT | None) -> _TSpacing:
        """
        Gets a copy of instance with below margin set or removed

        Args:
            value (float, UnitT, optional): Margin value (in ``mm`` units) or :ref:`proto_unit_obj`.

        Returns:
            Spacing: Indent instance
        """
        cp = self.copy()
        cp.prop_below = value
        return cp

    def fmt_style_no_space(self: _TSpacing, value: bool | None) -> _TSpacing:
        """
        Gets a copy of instance with style no spacing set or removed

        Args:
            value (bool | None): Auto value.

        Returns:
            Spacing: Indent instance
        """
        cp = self.copy()
        cp.prop_style_no_space = value
        return cp

    # endregion style methods

    # region Style Properties
    @property
    def style_no_space(self: _TSpacing) -> _TSpacing:
        """Gets copy of instance with style no spacing set"""
        cp = self.copy()
        cp.prop_style_no_space = True
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
    def prop_above(self) -> UnitMM | None:
        """Gets/Sets the top margin of the paragraph (in ``mm`` units)."""
        pv = cast(int, self._get("ParaTopMargin"))
        return None if pv is None else UnitMM.from_mm100(pv)

    @prop_above.setter
    def prop_above(self, value: float | UnitT | None):
        if value is None:
            self._remove("ParaTopMargin")
            return
        try:
            self._set("ParaTopMargin", value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set("ParaTopMargin", UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def prop_below(self) -> UnitMM | None:
        """Gets/Sets the bottom margin of the paragraph (in mm units)."""
        pv = cast(int, self._get("ParaBottomMargin"))
        return None if pv is None else UnitMM.from_mm100(pv)

    @prop_below.setter
    def prop_below(self, value: float | UnitT | None):
        if value is None:
            self._remove("ParaBottomMargin")
            return
        try:
            self._set("ParaBottomMargin", value.get_value_mm100())  # type: ignore
        except AttributeError:
            self._set("ParaBottomMargin", UnitConvert.convert_mm_mm100(value))  # type: ignore

    @property
    def prop_style_no_space(self) -> bool | None:
        """Gets/Sets if no space between paragraphs of the same style"""
        return self._get("ParaContextMargin")

    @prop_style_no_space.setter
    def prop_style_no_space(self, value: bool | None):
        if value is None:
            self._remove("ParaContextMargin")
            return
        self._set("ParaContextMargin", value)

    @property
    def default(self: _TSpacing) -> _TSpacing:
        """Gets ``Spacing`` default."""
        # pylint: disable=protected-access
        # pylint: disable=unexpected-keyword-arg
        try:
            return self._default_inst
        except AttributeError:
            self._default_inst = self.__class__(
                above=0.0, below=0.0, style_no_space=False, _cattribs=self._get_internal_cattribs()  # type: ignore
            )
            self._default_inst._is_default_inst = True
        return self._default_inst

    # endregion properties
