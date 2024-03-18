from __future__ import annotations
from typing import Any, TYPE_CHECKING
from ooodev.format.inner.kind.format_kind import FormatKind as FormatKind
from ooodev.mock.mock_g import DOCS_BUILDING

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from typing_extensions import Self
    from ooodev.units.unit_convert import UnitLength

    class UnitT(Protocol):
        """
        Protocol Class for units.

        .. seealso::

            :ref:`ns_units`

        .. _proto_unit_obj:

        UnitT
        =====

        """

        # if value is a attribute and not a property then value: Union[float, int] will not work.
        # a property can be a float only and still work. int is a subclass of float.
        # see: https://mypy.readthedocs.io/en/stable/common_issues.html#covariant-subtyping-of-mutable-protocol-members-is-rejected

        @property
        def value(self) -> float | int:
            """Unit actual value. Generally a ``float`` or ``int``"""
            ...

        @staticmethod
        def get_unit_length() -> UnitLength:
            """
            Gets instance unit length.

            Returns:
                UnitLength: Instance unit length ``UnitLength.PX``.

            .. versionadded:: 0.34.1
            """
            ...

        def convert_to(self, unit: UnitLength) -> float:
            """
            Converts instance value to specified unit.

            Args:
                unit (UnitLength): Unit to convert to.

            Returns:
                float: Value in specified unit.

            .. versionadded:: 0.34.1
            """
            ...

        def get_value_mm(self) -> float:
            """
            Gets instance value converted to Size in ``mm`` units.

            Returns:
                float: Value in ``mm`` units.
            """
            ...

        def get_value_mm100(self) -> int:
            """
            Gets instance value converted to Size in ``1/100th mm`` units.

            Returns:
                int: Value in ``1/100th mm`` units.
            """
            ...

        def get_value_pt(self) -> float:
            """
            Gets instance value converted to Size in ``pt`` (point) units.

            Returns:
                float: Value in ``pt`` units.
            """
            ...

        def get_value_px(self) -> float:
            """
            Gets instance value in ``px`` (pixel) units.

            Returns:
                float: Value in ``px`` units.
            """
            ...

        def get_value_app_font(self) -> float:
            """
            Gets instance value in ``AppFont`` units.

            Returns:
                float: Value in ``AppFont`` units.
            """
            ...

        @classmethod
        def from_unit_val(cls, value: UnitT | float | int) -> Self:
            """
            Get instance from ``UnitT`` or float or int value.

            Args:
                value (UnitT, float, int): ``UnitT`` or float value. If float then it is assumed to be in ``1/100th mm`` units.

            Returns:
                Self: Instance.

            .. versionadded:: 0.34.1
            """
            ...

        @classmethod
        def from_app_font(cls, value: float) -> Self:
            """
            Get instance from ``AppFont`` value.

            Args:
                value (int): ``AppFont`` value.

            Returns:
                Self: Instance.
            """
            ...

else:
    UnitT = Any

UnitObj = UnitT
