"""
Module for managing paragraph breaks.

.. seealso::

    - :ref:`help_writer_format_direct_para_outline_and_list`

.. versionadded:: 0.9.0
"""

from __future__ import annotations
from typing import Any, Tuple, overload, Type, TypeVar

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.common.props.list_style_props import ListStyleProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.format.writer.style.lst.style_list_kind import StyleListKind
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps

# from ...events.args.key_val_cancel_args import KeyValCancelArgs

_TListStyle = TypeVar("_TListStyle", bound="ListStyle")


class ListStyle(StyleBase):
    """
    Paragraph ListStyle

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. seealso::

        - :ref:`help_writer_format_direct_para_outline_and_list`

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(self, list_style: str | StyleListKind | None = None, num_start: int | None = None) -> None:
        """
        Constructor

        Args:
            list_style (str, StyleListKind, optional): List Style.
            num_start (int, optional): Starts with number.
                If ``-1`` then restart numbering at current paragraph is consider to be ``False``.
                If ``-2`` then restart numbering at current paragraph is consider to be ``True``.
                Otherwise, number starts at the value passed in.

        Returns:
            None:

        Note:
            If argument ``list_style`` is ``StyleListKind.NONE`` or empty string then ``num_start`` is ignored.

        See Also:

            - :ref:`help_writer_format_direct_para_outline_and_list`
        """
        # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1style_1_1ParagraphProperties-members.html

        init_vals = {}
        if list_style is not None:
            if str_style := str(list_style):
                init_vals[self._props.name] = str_style
            else:
                init_vals[self._props.name] = ""
                init_vals[self._props.value] = -1
                init_vals[self._props.restart] = False

        if num_start is not None and self._props.value not in init_vals:
            # ignore num_start if NumberingStartValue = -1 due to no style
            if num_start == -1:
                init_vals[self._props.value] = -1
                init_vals[self._props.restart] = False
            elif num_start < -1:
                init_vals[self._props.value] = -1
                init_vals[self._props.restart] = True
            elif num_start >= 0:
                init_vals[self._props.value] = num_start
                init_vals[self._props.restart] = True

        super().__init__(**init_vals)

    # endregion init

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        return ("com.sun.star.style.ParagraphProperties", "com.sun.star.style.ParagraphStyle")

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
        Applies break properties to ``obj``

        Args:
            obj (object): UNO object that supports ``com.sun.star.style.ParagraphProperties`` service.

        Returns:
            None:
        """
        super().apply(obj, **kwargs)

    def _props_set(self, obj: Any, **kwargs: Any) -> None:
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
    def from_obj(cls: Type[_TListStyle], obj: Any) -> _TListStyle: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TListStyle], obj: Any, **kwargs) -> _TListStyle: ...

    @classmethod
    def from_obj(cls: Type[_TListStyle], obj: Any, **kwargs) -> _TListStyle:
        """
        Gets instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            ListStyle: ``ListStyle`` instance that represents ``obj`` properties.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        def set_prop(key: str, o: ListStyle):
            nonlocal obj
            val = mProps.Props.get(obj, key, None)
            if val is not None:
                o._set(key, val)

        set_prop(inst._props.name, inst)
        set_prop(inst._props.value, inst)
        set_prop(inst._props.restart, inst)
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()

    # endregion methods

    # region Style Methods
    def fmt_list_style(self: _TListStyle, value: str | StyleListKind | None) -> _TListStyle:
        """
        Gets a copy of instance with before list style set or removed

        Args:
            value (str, StyleListKind, None): List style value.

        Returns:
            ListStyle: List Style instance
        """
        cp = self.copy()
        cp.prop_list_style = value
        return cp

    def fmt_num_start(self: _TListStyle, value: int | None) -> _TListStyle:
        """
        Gets a copy of instance with before list style set or removed

        Args:
            value (int | None): List style value.
                If ``-1`` then restart numbering at current paragraph is consider to be ``False``.
                If ``-2`` then restart numbering at current paragraph is consider to be ``True``.
                Otherwise, restart numbering is considered to be ``True``.

        Returns:
            ListStyle: List Style instance
        """
        cp = self.copy()
        cp.prop_num_start = value
        return cp

    # endregion Style Methods

    # region Style Properties
    @property
    def restart_numbers(self: _TListStyle) -> _TListStyle:
        """Gets instance with restart numbers set"""
        cp = self.copy()
        cp.prop_num_start = -1
        cp._set(self._props.restart, True)
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
    def prop_list_style(self) -> str | None:
        """Gets/Sets List Style"""
        return self._get(self._props.name)

    @prop_list_style.setter
    def prop_list_style(self, value: str | StyleListKind | None) -> None:
        if value is None:
            self._remove(self._props.name)
            return
        str_val = str(value)
        if not str_val:
            # empty string
            self._remove(self._props.name)
            return
        self._set(self._props.name, value)

    @property
    def prop_num_start(self) -> int | None:
        """
        Gets/Sets Starts with number.

        If Less then zero then restart numbering at current paragraph is consider to be ``False``;
        Otherwise; restart numbering is considered to be ``True``.
        """
        return self._get(self._props.value)

    @prop_num_start.setter
    def prop_num_start(self, value: int | None) -> None:
        if value is None:
            self._remove(self._props.value)
            self._remove(self._props.restart)
            return
        if value == -1:
            self._set(self._props.value, -1)
            self._set(self._props.restart, False)
            return
        if value < -1:
            self._set(self._props.value, -1)
            self._set(self._props.restart, True)
            return
        self._set(self._props.value, value)
        self._set(self._props.restart, True)

    @property
    def _props(self) -> ListStyleProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            # style_name="NumberingStyleName" is only used for default to reset to no list
            self._props_internal_attributes = ListStyleProps(
                name="NumberingStyleName", value="NumberingStartValue", restart="ParaIsNumberingRestart"
            )
        return self._props_internal_attributes

    @property
    def default(self: ListStyle) -> ListStyle:
        """Gets ``ListStyle`` default."""
        try:
            return self._default_inst
        except AttributeError:
            # empty string name will result in No List, list style being applied
            ls = ListStyle(list_style="")
            # as soon as name is set to empty string
            # _props.restart will automatically become false and cannot be change again
            # until a name is set.
            ls._is_default_inst = True
            self._default_inst = ls
        return self._default_inst

    # endregion properties
