from __future__ import annotations
from typing import Any, TYPE_CHECKING, Union
from ooodev.format.inner.kind.format_kind import FormatKind as FormatKind
from ooodev.mock.mock_g import DOCS_BUILDING

if TYPE_CHECKING or DOCS_BUILDING:
    from typing_extensions import Protocol
    from typing_extensions import Self
    from ooodev.units.unit_convert import UnitLength
    from ooodev.utils.kind.point_size_kind import PointSizeKind

    class UnitT(Protocol):
        """
        Protocol Class for units.

        .. seealso::

            :ref:`ns_units`
        """

        # if value is a attribute and not a property then value: Union[float, int] will not work.
        # a property can be a float only and still work. int is a subclass of float.
        # see: https://mypy.readthedocs.io/en/stable/common_issues.html#covariant-subtyping-of-mutable-protocol-members-is-rejected

        def __int__(self) -> int:
            """
            Gets instance value as an integer.

            Returns:
                int: Value as an integer.
            """
            ...

        @property
        def value(self) -> Union[float, int]:
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

        def get_value_app_font(self, kind: Union[PointSizeKind, int]) -> float:
            """
            Gets instance value in ``AppFont`` units.

            Returns:
                float: Value in ``AppFont`` units.
                kind (PointSizeKind, optional): The kind of ``AppFont`` to use.

            Note:
                AppFont units have different values when converted.
                This is true even if they have the same value in ``AppFont`` units.
                ``AppFontX(10)`` is not equal to ``AppFontY(10)`` when they are converted to different units.

                ``Kind`` when ``int`` is used, the value must be one of the following:

                - ``0`` is ``PointSizeKind.X``,
                - ``1`` is ``PointSizeKind.Y``,
                - ``2`` is ``PointSizeKind.WIDTH``,
                - ``3`` is ``PointSizeKind.HEIGHT``.

            Hint:
                - ``PointSizeKind`` can be imported from ``ooodev.utils.kind.point_size_kind``.
            """
            ...

        @classmethod
        def from_unit_val(cls, value: Union[UnitT, float, int]) -> Self:
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
        def from_app_font(cls, value: float, kind: Union[PointSizeKind, int]) -> Self:
            """
            Get instance from ``AppFont`` value.

            Args:
                value (int): ``AppFont`` value.
                kind (PointSizeKind): The kind of ``AppFont`` to use.

            Returns:
                UnitPX:

            Note:
                AppFont units have different values when converted.
                This is true even if they have the same value in ``AppFont`` units.
                ``AppFontX(10)`` is not equal to ``AppFontY(10)`` when they are converted to different units.

                ``Kind`` when ``int`` is used, the value must be one of the following:

                - ``0`` is ``PointSizeKind.X``,
                - ``1`` is ``PointSizeKind.Y``,
                - ``2`` is ``PointSizeKind.WIDTH``,
                - ``3`` is ``PointSizeKind.HEIGHT``.

            Hint:
                - ``PointSizeKind`` can be imported from ``ooodev.utils.kind.point_size_kind``.
            """
            ...

else:
    UnitT = Any

UnitObj = UnitT
