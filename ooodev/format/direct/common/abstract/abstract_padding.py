"""
Modele for managing paragraph padding.

.. versionadded:: 0.9.0
"""
from __future__ import annotations
from typing import Tuple, cast, overload, Type, TypeVar

from .....events.args.cancel_event_args import CancelEventArgs
from .....utils import props as mProps
from .....exceptions import ex as mEx
from .....utils import lo as mLo
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase
from ..props.border_props import BorderProps as BorderProps

_TAbstractPadding = TypeVar(name="_TAbstractPadding", bound="AbstractPadding")


class AbstractPadding(StyleBase):
    """
    Paragraph Padding

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        *,
        left: float | None = None,
        right: float | None = None,
        top: float | None = None,
        bottom: float | None = None,
        padding_all: float | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (float, optional): Paragraph left padding (in mm units).
            right (float, optional): Paragraph right padding (in mm units).
            top (float, optional): Paragraph top padding (in mm units).
            bottom (float, optional): Paragraph bottom padding (in mm units).
            padding_all (float, optional): Paragraph left, right, top, bottom padding (in mm units). If argument is present then ``left``, ``right``, ``top``, and ``bottom`` arguments are ignored.

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
        if padding_all is None:
            for key, value in zip(self._props, (left, top, right, bottom)):
                set_val(key, value)
        else:
            for key in self._props:
                set_val(key, padding_all)

        super().__init__(**init_vals)

    # endregion init

    # region methods
    def _on_modifing(self, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifing(event)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.style.ParagraphProperties",)
        return self._supported_services_values

    # region apply()
    @overload
    def apply(self, obj: object) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies padding to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.
            kwargs (Any, optional): Expandable list of key value pairs that may be used in child classes.

        Returns:
            None:
        """
        try:
            super().apply(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")
        return None

    # endregion apply()

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TAbstractPadding], obj: object) -> _TAbstractPadding:
        ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TAbstractPadding], obj: object, **kwargs) -> _TAbstractPadding:
        ...

    @classmethod
    def from_obj(cls: Type[_TAbstractPadding], obj: object, **kwargs) -> _TAbstractPadding:
        """
        Gets Border Padding instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            BorderPadding: BorderPadding that represents ``obj`` padding.
        """
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        inst._set(inst._props.left, int(mProps.Props.get(obj, inst._props.left)))
        inst._set(inst._props.right, int(mProps.Props.get(obj, inst._props.right)))
        inst._set(inst._props.top, int(mProps.Props.get(obj, inst._props.top)))
        inst._set(inst._props.bottom, int(mProps.Props.get(obj, inst._props.bottom)))

        return inst

    # endregion from_obj()
    # endregion methods

    # region style methods
    def fmt_padding_all(self: _TAbstractPadding, value: float | None) -> _TAbstractPadding:
        """
        Gets copy of instance with left, right, top, bottom sides set or removed

        Args:
            value (float | None): Padding value

        Returns:
            Padding: Padding instance
        """
        cp = self.copy()
        cp.prop_top = value
        cp.prop_bottom = value
        cp.prop_left = value
        cp.prop_right = value
        return cp

    def fmt_top(self: _TAbstractPadding, value: float | None) -> _TAbstractPadding:
        """
        Gets a copy of instance with top side set or removed

        Args:
            value (float | None): Padding value

        Returns:
            Padding: Padding instance
        """
        cp = self.copy()
        cp.prop_top = value
        return cp

    def fmt_bottom(self: _TAbstractPadding, value: float | None) -> _TAbstractPadding:
        """
        Gets a copy of instance with bottom side set or removed

        Args:
            value (float | None): Padding value

        Returns:
            Padding: Padding instance
        """
        cp = self.copy()
        cp.prop_bottom = value
        return cp

    def fmt_left(self: _TAbstractPadding, value: float | None) -> _TAbstractPadding:
        """
        Gets a copy of instance with left side set or removed

        Args:
            value (float | None): Padding value

        Returns:
            Padding: Padding instance
        """
        cp = self.copy()
        cp.prop_left = value
        return cp

    def fmt_right(self: _TAbstractPadding, value: float | None) -> _TAbstractPadding:
        """
        Gets a copy of instance with right side set or removed

        Args:
            value (float | None): Padding value

        Returns:
            Padding: Padding instance
        """
        cp = self.copy()
        cp.prop_right = value
        return cp

    # endregion style methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._fromat_kind_prop
        except AttributeError:
            self._fromat_kind_prop = FormatKind.PARA
        return self._fromat_kind_prop

    @property
    def prop_left(self) -> float | None:
        """Gets/Sets paragraph left padding (in mm units)."""
        pv = cast(int, self._get(self._props.left))
        if pv is None:
            return None
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_left.setter
    def prop_left(self, value: float | None):
        if value is None:
            self._remove(self._props.left)
            return
        self._set(self._props.left, round(value * 100))

    @property
    def prop_right(self) -> float | None:
        """Gets/Sets paragraph right padding (in mm units)."""
        pv = cast(int, self._get(self._props.right))
        if pv is None:
            return None
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_right.setter
    def prop_right(self, value: float | None):
        if value is None:
            self._remove(self._props.right)
            return
        self._set(self._props.right, round(value * 100))

    @property
    def prop_top(self) -> float | None:
        """Gets/Sets paragraph top padding (in mm units)."""
        pv = cast(int, self._get(self._props.top))
        if pv is None:
            return None
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_top.setter
    def prop_top(self, value: float | None):
        if value is None:
            self._remove(self._props.top)
            return
        self._set(self._props.top, round(value * 100))

    @property
    def prop_bottom(self) -> float | None:
        """Gets/Sets paragraph bottom padding (in mm units)."""
        pv = cast(int, self._get(self._props.bottom))
        if pv is None:
            return None
        if pv == 0:
            return 0.0
        return float(pv / 100)

    @prop_bottom.setter
    def prop_bottom(self, value: float | None):
        if value is None:
            self._remove(self._props.bottom)
            return
        self._set(self._props.bottom, round(value * 100))

    @property
    def _props(self) -> BorderProps:
        try:
            return self.__border_properties
        except AttributeError:
            self.__border_properties = BorderProps(
                left="ParaLeftMargin", top="ParaTopMargin", right="ParaRightMargin", bottom="ParaBottomMargin"
            )
        return self.__border_properties

    # endregion properties
