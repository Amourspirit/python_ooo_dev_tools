from __future__ import annotations
from typing import Any, TYPE_CHECKING
from ooodev.format.inner.kind.format_kind import FormatKind as FormatKind

if TYPE_CHECKING:
    from typing_extensions import Protocol

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

else:
    UnitT = Any

UnitObj = UnitT
