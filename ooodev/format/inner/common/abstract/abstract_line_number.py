# region Import
from __future__ import annotations
from typing import NamedTuple
from typing import Any, Tuple, overload, Type, TypeVar

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase

# endregion Import

_TAbstractLineNumber = TypeVar("_TAbstractLineNumber", bound="AbstractLineNumber")


class LineNumberProps(NamedTuple):
    value: str
    count: str


class AbstractLineNumber(StyleBase):
    """
    Paragraph Line Numbers

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(self, num_start: int = 0) -> None:
        """
        Constructor

        Args:
            num_start (int, optional): Restart paragraph with number.
                If ``0`` then this paragraph is include in line numbering.
                If ``-1`` then this paragraph is excluded in line numbering.
                If greater than zero this paragraph is included in line numbering and the numbering is restarted with
                value of ``num_start``.

        Returns:
            None:
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html

        init_vals = {}

        if num_start == 0:
            init_vals[self._props.value] = 0
            init_vals[self._props.count] = True
        elif num_start < 0:
            init_vals[self._props.value] = 0
            init_vals[self._props.count] = False
        else:
            init_vals[self._props.value] = num_start
            init_vals[self._props.count] = True

        super().__init__(**init_vals)

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
    def apply(self, obj: Any) -> None:  # type: ignore
        ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies break properties to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            None:
        """
        super().apply(obj, **kwargs)

    def _props_set(self, obj: object, **kwargs: Any) -> None:
        try:
            super()._props_set(obj, **kwargs)
        except mEx.MultiError as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply(): Unable to set Property")
            for err in e.errors:
                mLo.Lo.print(f"  {err}")

    # endregion apply()
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TAbstractLineNumber], obj: object) -> _TAbstractLineNumber: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TAbstractLineNumber], obj: object, **kwargs) -> _TAbstractLineNumber: ...

    @classmethod
    def from_obj(cls: Type[_TAbstractLineNumber], obj: object, **kwargs) -> _TAbstractLineNumber:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            LineNum: ``LineNum`` instance that represents ``obj`` properties.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        def set_prop(key: str, o: AbstractLineNumber):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if val is not None:
                o._set(key, val)

        set_prop(inst._props.value, inst)
        set_prop(inst._props.count, inst)
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()

    # endregion methods

    # region Style Methods

    def fmt_num_start(self: _TAbstractLineNumber, value: int) -> _TAbstractLineNumber:
        """
        Gets a copy of instance with before list style set or removed

        Args:
            value (int): List style value.
                If ``0`` then this paragraph is include in line numbering.
                If ``-1`` then this paragraph is excluded in line numbering.
                If greater than zero this paragraph is included in line numbering and the numbering is restarted
                with ``value``.

        Returns:
            LineNum: Line Number instance
        """
        cp = self.copy()
        cp.prop_num_start = value
        return cp

    # endregion Style Methods

    # region Style Properties
    @property
    def restart_numbers(self: _TAbstractLineNumber) -> _TAbstractLineNumber:
        """Gets instance with restart numbers set to ``1``"""
        return self.__class__(1)

    @property
    def include(self: _TAbstractLineNumber) -> _TAbstractLineNumber:
        """Gets instance with include in line numbering set to include."""
        cp = self.copy()
        # zero or higher is already include
        if cp.prop_num_start is not None:
            cp.prop_num_start = max(cp.prop_num_start, 0)
        return cp

    @property
    def exclude(self: _TAbstractLineNumber) -> _TAbstractLineNumber:
        """Gets instance with include in line numbering set to exclude."""
        return self.__class__(-1)

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
    def prop_num_start(self) -> int | None:
        """
        Gets/Sets Restart at this paragraph number.

        If Less than zero restart numbering at current paragraph is consider to be ``False``;
        Otherwise; restart numbering is considered to be ``True``.
        """
        return self._get(self._props.value)

    @prop_num_start.setter
    def prop_num_start(self, value: int) -> None:
        if value == 0:
            self._set(self._props.value, 0)
            self._set(self._props.count, True)
            return
        if value < 0:
            self._set(self._props.value, 0)
            self._set(self._props.count, False)
        self._set(self._props.value, value)
        self._set(self._props.count, True)

    @property
    def _props(self) -> LineNumberProps:
        try:
            return self._props_line_num
        except AttributeError:
            self._props_line_num = LineNumberProps(value="ParaLineNumberStartValue", count="ParaLineNumberCount")
        return self._props_line_num

    @property
    def default(self: _TAbstractLineNumber) -> _TAbstractLineNumber:
        """Gets ``LineNum`` default."""
        try:
            return self._default_inst
        except AttributeError:
            self._default_inst = self.__class__(num_start=0, _cattribs=self._get_internal_cattribs())
            self._default_inst._is_default_inst = True
        return self._default_inst

    # endregion properties
