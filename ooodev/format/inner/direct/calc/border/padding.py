"""
Module for managing paragraph padding.

.. versionadded:: 0.9.0
"""

# region Import
from __future__ import annotations
from typing import Any, Type, TypeVar, overload

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.units.unit_obj import UnitT
from ooodev.utils import props as mProps
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.common.abstract.abstract_padding import AbstractPadding
from ooodev.format.inner.common.props.border_props import BorderProps as BorderProps

# endregion Import

_TPadding = TypeVar(name="_TPadding", bound="Padding")


class Padding(AbstractPadding):
    """
    Paragraph Padding

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. seealso::

        - :ref:`help_calc_format_direct_cell_borders`
        - :ref:`help_writer_format_direct_table_borders`

    .. versionadded:: 0.9.0
    """

    def __init__(
        self,
        *,
        left: float | UnitT | None = None,
        right: float | UnitT | None = None,
        top: float | UnitT | None = None,
        bottom: float | UnitT | None = None,
        all: float | UnitT | None = None,
    ) -> None:
        """
        Constructor

        Args:
            left (float, UnitT, optional): Left (in ``mm`` units) or :ref:`proto_unit_obj`.
            right (float, UnitT, optional): Right (in ``mm`` units)  or :ref:`proto_unit_obj`.
            top (float, UnitT, optional): Top (in ``mm`` units)  or :ref:`proto_unit_obj`.
            bottom (float, UnitT,  optional): Bottom (in ``mm`` units)  or :ref:`proto_unit_obj`.
            all (float, UnitT, optional): Left, right, top, bottom (in ``mm`` units)  or :ref:`proto_unit_obj`.
                If argument is present then ``left``, ``right``, ``top``, and ``bottom`` arguments are ignored.

        Raises:
            ValueError: If any argument value is less than zero.

        Returns:
            None:

        See Also:

            - :ref:`help_calc_format_direct_cell_borders`
            - :ref:`help_writer_format_direct_table_borders`
        """
        super().__init__(left=left, right=right, top=top, bottom=bottom, all=all)

    # region methods

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
            obj (object): UNO Object.

        Raises:
            NotSupportedServiceError: If ``obj`` is not supported.

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
        return inst

    # endregion from_obj()
    def _on_modifying(self, source: Any, event: CancelEventArgs) -> None:
        if self._is_default_inst:
            raise ValueError("Modifying a default instance is not allowed")
        return super()._on_modifying(source, event)

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

    @property
    def _props(self) -> BorderProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = BorderProps(
                left="ParaLeftMargin", top="ParaTopMargin", right="ParaRightMargin", bottom="ParaBottomMargin"
            )
        return self._props_internal_attributes

    # endregion properties
