"""
Module for managing character padding.

.. versionadded:: 0.9.0
"""

# region Import
from __future__ import annotations
from typing import Tuple, TypeVar, TYPE_CHECKING

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.common.abstract.abstract_padding import AbstractPadding
from ooodev.format.inner.common.props.border_props import BorderProps

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT

# endregion Import

_TPadding = TypeVar(name="_TPadding", bound="Padding")


class Padding(AbstractPadding):
    """
    Paragraph Border Padding

    Any properties starting with ``prop_`` set or get current instance values.

    All methods starting with ``fmt_`` can be used to chain together properties.

    .. seealso::

        - :ref:`help_writer_format_direct_char_borders`

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

            - :ref:`help_writer_format_direct_char_borders`
        """
        super().__init__(left=left, right=right, top=top, bottom=bottom, all=all)

    # region methods
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.style.CharacterProperties",
                "com.sun.star.style.CharacterStyle",
            )
        return self._supported_services_values

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
                left="CharLeftBorderDistance",
                top="CharTopBorderDistance",
                right="CharRightBorderDistance",
                bottom="CharBottomBorderDistance",
            )
        return self._props_internal_attributes

    @property
    def default(self: _TPadding) -> _TPadding:  # type: ignore[misc]
        """Gets BorderPadding default."""
        # pylint: disable=unexpected-keyword-arg
        # pylint: disable=protected-access
        try:
            return self._default_inst
        except AttributeError:
            inst = self.__class__(all=0.0, _cattribs=self._get_internal_cattribs())  # type: ignore
            inst._is_default_inst = True
            self._default_inst = inst
        return self._default_inst

    # endregion properties
