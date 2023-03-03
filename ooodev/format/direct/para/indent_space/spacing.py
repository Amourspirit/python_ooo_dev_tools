"""
Modele for managing paragraph padding.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, cast, overload, Type, TypeVar

from .....events.args.cancel_event_args import CancelEventArgs
from .....exceptions import ex as mEx
from .....meta.static_prop import static_prop
from .....proto.unit_obj import UnitObj
from .....utils import lo as mLo
from .....utils import props as mProps
from .....utils.data_type.unit_mm import UnitMM
from .....utils.unit_convert import UnitConvert
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase

_TSpacing = TypeVar(name="_TSpacing", bound="Spacing")


class Spacing(StyleBase):
    """
    Paragraph Spacing

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        *,
        above: float | UnitObj | None = None,
        below: float | UnitObj | None = None,
        style_no_space: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            above (float, UnitObj, optional): Determines the top margin of the paragraph (in ``mm`` units) or :ref:`proto_unit_obj`.
            below (float, UnitObj, optional): Determines the bottom margin of the paragraph (in ``mm`` units) or :ref:`proto_unit_obj`.
            style_no_space (bool, optional): Do not add space between paragraphs of the same style.
        Returns:
            None:
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html
        super().__init__()

        if not above is None:
            self.prop_above = above
        if not below is None:
            self.prop_below = below
        if not style_no_space is None:
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

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TSpacing], obj: object) -> _TSpacing:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TSpacing], obj: object, **kwargs) -> _TSpacing:
        ...

    @classmethod
    def from_obj(cls: Type[_TSpacing], obj: object, **kwargs) -> _TSpacing:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Spacing: ``Spacing`` instance that represents ``obj`` spacing.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        def set_prop(key: str, indent: Spacing):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if not val is None:
                indent._set(key, val)

        set_prop("ParaTopMargin", inst)
        set_prop("ParaBottomMargin", inst)
        set_prop("ParaContextMargin", inst)
        return inst

    # endregion from_obj()

    # endregion methods

    # region style methods
    def fmt_above(self: _TSpacing, value: float | UnitObj | None) -> _TSpacing:
        """
        Gets a copy of instance with above margin set or removed

        Args:
            value (float, UnitObj, optional): Margin value (in ``mm`` units) or :ref:`proto_unit_obj`.

        Returns:
            Spacing: Indent instance
        """
        cp = self.copy()
        cp.prop_above = value
        return cp

    def fmt_below(self: _TSpacing, value: float | UnitObj | None) -> _TSpacing:
        """
        Gets a copy of instance with below margin set or removed

        Args:
            value (float, UnitObj, optional): Margin value (in ``mm`` units) or :ref:`proto_unit_obj`.

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
        if pv is None:
            return None
        return UnitMM.from_mm100(pv)

    @prop_above.setter
    def prop_above(self, value: float | UnitObj | None):
        if value is None:
            self._remove("ParaTopMargin")
            return
        try:
            self._set("ParaTopMargin", value.get_value_mm100())
        except AttributeError:
            self._set("ParaTopMargin", UnitConvert.convert_mm_mm100(value))

    @property
    def prop_below(self) -> UnitMM | None:
        """Gets/Sets the bottom margin of the paragraph (in mm units)."""
        pv = cast(int, self._get("ParaBottomMargin"))
        if pv is None:
            return None
        return UnitMM.from_mm100(pv)

    @prop_below.setter
    def prop_below(self, value: float | UnitObj | None):
        if value is None:
            self._remove("ParaBottomMargin")
            return
        try:
            self._set("ParaBottomMargin", value.get_value_mm100())
        except AttributeError:
            self._set("ParaBottomMargin", UnitConvert.convert_mm_mm100(value))

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

    @static_prop
    def default() -> Spacing:  # type: ignore[misc]
        """Gets ``Spacing`` default. Static Property."""
        try:
            return Spacing._DEFAULT_INST
        except AttributeError:
            Spacing._DEFAULT_INST = Spacing(above=0.0, below=0.0, style_no_space=False)
            Spacing._DEFAULT_INST._is_default_inst = True
        return Spacing._DEFAULT_INST

    # endregion properties
