from __future__ import annotations
from typing import NamedTuple, cast
from typing import Any, Tuple, overload, Type, TypeVar

from .....events.args.cancel_event_args import CancelEventArgs
from .....exceptions import ex as mEx
from .....utils import lo as mLo
from .....utils import props as mProps
from ....kind.format_kind import FormatKind
from ....style_base import StyleBase
from ..props.hf_props import HfProps
from .....utils.unit_convert import UnitConvert

# from ...events.args.key_val_cancel_args import KeyValCancelArgs

_TAbstractHF = TypeVar(name="_TAbstractHF", bound="AbstractHF")


class AbstractHF(StyleBase):
    """
    Paragraph Line Numbers

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        *,
        on: bool | None = None,
        shared: bool | None = None,
        shared_first: bool | None = None,
        margin_left: float | None = None,
        margin_right: float | None = None,
        spacing: float | None = None,
        spacing_dyn: bool | None = None,
        height: float | None = None,
        height_auto: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            on (bool | None, optional): Specifices if section is on.
            shared (bool | None, optional): Specifies if same contents left and right.
            shared_first (bool | None, optional): Specifies if same contents on first page.
            margin_left (float | None, optional): Specifies Left Margin in ``mm`` units.
            margin_right (float | None, optional): Specifies Right Margin in ``mm`` units.
            spacing (float | None, optional): Specifies Spacing in ``mm`` units.
            spacing_dyn (bool | None, optional): Specifies if if dynamic spacing is used.
            height (float | None, optional): Specifies Height in ``mm`` units.
            height_auto (bool | None, optional): Specifies if auto fit height is used.
        """

        super().__init__()
        self.prop_on = on
        self.prop_shared = shared
        self.prop_shared_first = shared_first
        self.prop_margin_left = margin_left
        self.prop_margin_right = margin_right
        self.prop_spacing = spacing
        self.prop_spacing_dynamic = spacing_dyn
        self.prop_height = height
        self.prop_height_auto = height_auto

    # endregion init

    # region methods

    # region Internal Methods
    def _get_spacing(self) -> int:
        val = cast(int, self._get(self._props.spacing))
        if val is None:
            return 0
        return val

    # endregion Internal Methods

    # region Overrides
    def _supported_services(self) -> Tuple[str, ...]:
        return ("com.sun.star.style.ParagraphProperties", "com.sun.star.style.ParagraphStyle")

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
        Applies break properties to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            None:
        """
        super().apply(obj, **kwargs)

    # endregion apply()

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion Overrides

    # region Static Methods
    @classmethod
    def from_obj(cls: Type[_TAbstractHF], obj: object) -> _TAbstractHF:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            LineNum: ``LineNum`` instance that represents ``obj`` properties.
        """
        inst = super(AbstractHF, cls).__new__(cls)
        inst.__init__()
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        def set_prop(key: str, o: AbstractHF):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if not val is None:
                o._set(key, val)

        set_prop(inst._props.on, inst)
        set_prop(inst._props.shared, inst)
        set_prop(inst._props.shared_first, inst)
        set_prop(inst._props.margin_left, inst)
        set_prop(inst._props.margin_right, inst)
        set_prop(inst._props.spacing, inst)
        set_prop(inst._props.spacing_dyn, inst)
        set_prop(inst._props.height, inst)
        set_prop(inst._props.height_auto, inst)

        return inst

    # endregion Static Methods

    # endregion methods

    # region Style Methods

    def fmt_on(self: _TAbstractHF, value: bool | None) -> _TAbstractHF:
        """
        Gets a copy of instance with ``on`` set or removed.

        Args:
            value (bool | None): Specifices if section is on.

        Returns:
            AbstractHF:
        """
        cp = self.copy()
        cp.prop_on = value
        return cp

    def fmt_shared(self: _TAbstractHF, value: bool | None) -> _TAbstractHF:
        """
        Gets a copy of instance with ``shared`` set or removed.

        Args:
            value (bool | None): Specifies if same contents left and right.

        Returns:
            AbstractHF:
        """
        cp = self.copy()
        cp.prop_shared = value
        return cp

    def fmt_shared_first(self: _TAbstractHF, value: bool | None) -> _TAbstractHF:
        """
        Gets a copy of instance with ``shared_first`` set or removed.

        Args:
            value (bool | None): Specifies if same contents on first page.

        Returns:
            AbstractHF:
        """
        cp = self.copy()
        cp.prop_shared_first = value
        return cp

    def fmt_margin_left(self: _TAbstractHF, value: float | None) -> _TAbstractHF:
        """
        Gets a copy of instance with ``margin_left`` set or removed.

        Args:
            value (float | None): Specifies Left Margin in ``mm`` units.

        Returns:
            AbstractHF:
        """
        cp = self.copy()
        cp.prop_margin_right = value
        return cp

    def fmt_margin_right(self: _TAbstractHF, value: float | None) -> _TAbstractHF:
        """
        Gets a copy of instance with ``margin_right`` set or removed.

        Args:
            value (float | None): Specifies Right Margin in ``mm`` units.

        Returns:
            AbstractHF:
        """
        cp = self.copy()
        cp.prop_margin_left = value
        return cp

    def fmt_spacing(self: _TAbstractHF, value: float | None) -> _TAbstractHF:
        """
        Gets a copy of instance with ``spacing`` set or removed.

        Args:
            value (float | None): Specifies Spacing in ``mm`` units.

        Returns:
            AbstractHF:
        """
        cp = self.copy()
        cp.prop_spacing = value
        return cp

    def fmt_spacing_dyn(self: _TAbstractHF, value: bool | None) -> _TAbstractHF:
        """
        Gets a copy of instance with Dynamic Spacing set or removed.

        Args:
            value (bool | None): Specifies if if dynamic spacing is used.

        Returns:
            AbstractHF:
        """
        cp = self.copy()
        cp.prop_spacing_dynamic = value
        return cp

    def fmt_height(self: _TAbstractHF, value: float | None) -> _TAbstractHF:
        """
        Gets a copy of instance with ``height`` set or removed.

        Args:
            value (float | None): Specifies Height in ``mm`` units.

        Returns:
            AbstractHF:
        """
        cp = self.copy()
        cp.prop_height = value
        return cp

    def fmt_height_aut(self: _TAbstractHF, value: bool | None) -> _TAbstractHF:
        """
        Gets a copy of instance with Dynamic Height set or removed.

        Args:
            value (bool | None): Specifies if auto fit height is used.

        Returns:
            AbstractHF:
        """
        cp = self.copy()
        cp.prop_height_auto = value
        return cp

    # endregion Style Methods

    # region Style Properties
    @property
    def restart_on(self: _TAbstractHF) -> _TAbstractHF:
        """Gets instance section ``on`` set to ``True``"""
        cp = self.copy()
        cp.prop_on = True
        return cp

    @property
    def shared(self: _TAbstractHF) -> _TAbstractHF:
        """Gets instance ``shared``  set to ``True``"""
        cp = self.copy()
        cp.prop_shared = True
        return cp

    @property
    def space_dyn(self: _TAbstractHF) -> _TAbstractHF:
        """Gets instance Dynamic spacings  set to ``True``"""
        cp = self.copy()
        cp.prop_spacing_dynamic = True
        return cp

    @property
    def height_auto(self: _TAbstractHF) -> _TAbstractHF:
        """Gets instance ``height_auto``  set to ``True``"""
        cp = self.copy()
        cp.prop_height_auto = True
        return cp

    # endregion Style Properties

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        return FormatKind.PARA

    @property
    def prop_on(self) -> bool | None:
        """
        Gets/Sets if section is on.
        """
        return self._get(self._props.on)

    @prop_on.setter
    def prop_on(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.on)
            return
        self._set(self._props.on, value)

    @property
    def prop_shared(self) -> bool | None:
        """
        Gets/Sets if same contents left and right.
        """
        return self._get(self._props.shared)

    @prop_shared.setter
    def prop_shared(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.shared)
            return
        self._set(self._props.shared, value)

    @property
    def prop_shared_first(self) -> bool | None:
        """
        Gets/Sets if same contents on first page.
        """
        return self._get(self._props.shared_first)

    @prop_shared_first.setter
    def prop_shared_first(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.shared_first)
            return
        self._set(self._props.shared_first, value)

    @property
    def prop_margin_left(self) -> float | None:
        """
        Gets/Sets Left Margin in ``mm`` units.
        """
        pv = cast(int, self._get(self._props.margin_left))
        if pv is None:
            return None
        return UnitConvert.convert_mm100_mm(pv)

    @prop_margin_left.setter
    def prop_margin_left(self, value: float | None) -> None:
        if value is None:
            self._remove(self._props.margin_left)
            return
        self._set(self._props.margin_left, UnitConvert.convert_mm_mm100(value))

    @property
    def prop_margin_right(self) -> float | None:
        """
        Gets/Sets Right Margin in ``mm`` units.
        """
        pv = cast(int, self._get(self._props.margin_right))
        if pv is None:
            return None
        return UnitConvert.convert_mm100_mm(pv)

    @prop_margin_right.setter
    def prop_margin_right(self, value: float | None) -> None:
        if value is None:
            self._remove(self._props.margin_right)
            return
        self._set(self._props.margin_right, UnitConvert.convert_mm_mm100(value))

    @property
    def prop_spacing(self) -> float | None:
        """
        Gets/Sets Spacing in ``mm`` units.
        """
        pv = cast(int, self._get(self._props.spacing))
        if pv is None:
            return None
        return UnitConvert.convert_mm100_mm(pv)

    @prop_spacing.setter
    def prop_spacing(self, value: float | None) -> None:
        if value is None:
            self._remove(self._props.spacing)
            return
        self._set(self._props.spacing, UnitConvert.convert_mm_mm100(value))

    @property
    def prop_spacing_dynamic(self) -> bool | None:
        """
        Gets/Sets if dynamic spacing is used
        """
        return self._get(self._props.spacing_dyn)

    @prop_spacing_dynamic.setter
    def prop_spacing_dynamic(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.spacing_dyn)
            return
        self._set(self._props.spacing_dyn, value)

    @property
    def prop_height(self) -> float | None:
        """
        Gets/Sets Height in ``mm`` units.
        """
        # height includes spacing
        pv = cast(int, self._get(self._props.height))
        if pv is None:
            return None
        if pv is 0:
            return 0
        spacing = self._get_spacing()
        val = pv - spacing
        return UnitConvert.convert_mm100_mm(val)

    @prop_height.setter
    def prop_height(self, value: float | None) -> None:
        # height includes spacing
        if value is None:
            self._remove(self._props.height)
            return
        spacing = self._get_spacing()
        self._set(self._props.height, UnitConvert.convert_mm_mm100(value) + spacing)

    @property
    def prop_height_auto(self) -> bool | None:
        """
        Gets/Sets if auto fit height is used
        """
        return self._get(self._props.height_auto)

    @prop_height_auto.setter
    def prop_height_auto(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.height_auto)
            return
        self._set(self._props.height_auto, value)

    @property
    def _props(self) -> HfProps:
        raise NotImplementedError

    # endregion properties
