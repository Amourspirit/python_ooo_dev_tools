"""
Module for managing paragraph padding.

.. versionadded:: 0.9.0
"""

# region Imports
from __future__ import annotations
from typing import Any, Tuple, Type, TypeVar, overload

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.common.abstract.abstract_padding import AbstractPadding

# endregion Imports

_TPadding = TypeVar(name="_TPadding", bound="Padding")


class Padding(AbstractPadding):
    """
    Paragraph Padding

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.
    """

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

    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TPadding], obj: Any) -> _TPadding: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TPadding], obj: Any, **kwargs) -> _TPadding: ...

    @classmethod
    def from_obj(cls: Type[_TPadding], obj: Any, **kwargs) -> _TPadding:
        """
        Gets Padding instance from object

        Args:
            obj (object): UNO object.

        Raises:
            NotSupportedError: If ``obj`` is not supported.

        Returns:
            Padding: Padding that represents ``obj`` padding.
        """
        # pylint: disable=protected-access
        inst = cls(**kwargs)
        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError(f'Object is not supported for conversion to "{cls.__name__}"')

        inst._set(inst._props.left, int(mProps.Props.get(obj, inst._props.left)))
        inst._set(inst._props.right, int(mProps.Props.get(obj, inst._props.right)))
        inst._set(inst._props.top, int(mProps.Props.get(obj, inst._props.top)))
        inst._set(inst._props.bottom, int(mProps.Props.get(obj, inst._props.bottom)))
        inst.set_update_obj(obj)
        return inst

    # endregion from_obj()

    # endregion methods

    # region properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.PARA
        return self._format_kind_prop

    # endregion properties
