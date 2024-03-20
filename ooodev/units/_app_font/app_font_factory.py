from __future__ import annotations

from typing import TYPE_CHECKING, Type
from ooodev.utils.kind.point_size_kind import PointSizeKind

if TYPE_CHECKING:
    from ooodev.units._app_font.unit_app_font_base import UnitAppFontBase
    from ooodev.units.unit_app_font_height import UnitAppFontHeight
    from ooodev.units.unit_app_font_width import UnitAppFontWidth
    from ooodev.units.unit_app_font_x import UnitAppFontX
    from ooodev.units.unit_app_font_y import UnitAppFontY


class AppFontFactory:
    """
    Factory to create ``UnitAppFontBase`` units.

    This factory is used to create the various ``UnitAppFontBase`` units.
    """

    @staticmethod
    def get_app_font(kind: PointSizeKind | int, val: float) -> UnitAppFontBase:
        """
        Get the application font unit.

        Args:
            kind (PointSizeKind | int): The kind of the unit.


        Returns:
            UnitAppFontBase: The application font unit.

        Note:
            ``Kind`` when ``int`` is used, the value must be one of the following:

            - ``0`` is ``PointSizeKind.X``,
            - ``1`` is ``PointSizeKind.Y``,
            - ``2`` is ``PointSizeKind.WIDTH``,
            - ``3`` is ``PointSizeKind.HEIGHT``.
        """
        af_kind = PointSizeKind(kind)
        if af_kind is PointSizeKind.X:
            return AppFontFactory.create_app_font_x(val)
        if af_kind is PointSizeKind.Y:
            return AppFontFactory.create_app_font_y(val)
        if af_kind is PointSizeKind.WIDTH:
            return AppFontFactory.create_app_font_width(val)
        if af_kind is PointSizeKind.HEIGHT:
            return AppFontFactory.create_app_font_height(val)
        raise ValueError(f"Invalid kind: {kind}")

    @staticmethod
    def get_app_font_type(kind: PointSizeKind | int) -> Type[UnitAppFontBase]:
        """
        Get the application font unit type.

        Args:
            kind (PointSizeKind | int): The kind of the unit.

        Returns:
            Type[UnitAppFontBase]: The application font unit type.

        Note:
            ``Kind`` when ``int`` is used, the value must be one of the following:

            - ``0`` is ``PointSizeKind.X``,
            - ``1`` is ``PointSizeKind.Y``,
            - ``2`` is ``PointSizeKind.WIDTH``,
            - ``3`` is ``PointSizeKind.HEIGHT``.
        """
        af_kind = PointSizeKind(kind)
        if af_kind is PointSizeKind.X:
            from ooodev.units.unit_app_font_x import UnitAppFontX

            return UnitAppFontX
        if af_kind is PointSizeKind.Y:
            from ooodev.units.unit_app_font_y import UnitAppFontY

            return UnitAppFontY
        if af_kind is PointSizeKind.WIDTH:
            from ooodev.units.unit_app_font_width import UnitAppFontWidth

            return UnitAppFontWidth
        if af_kind is PointSizeKind.HEIGHT:
            from ooodev.units.unit_app_font_height import UnitAppFontHeight

            return UnitAppFontHeight
        raise ValueError(f"Invalid kind: {kind}")

    @staticmethod
    def create_app_font_height(val: float) -> UnitAppFontHeight:
        """
        Create a ``UnitAppFontHeight`` unit.

        Returns:
            UnitAppFontHeight: The created unit.
        """
        from ooodev.units.unit_app_font_height import UnitAppFontHeight

        return UnitAppFontHeight(val)

    @staticmethod
    def create_app_font_width(val: float) -> UnitAppFontWidth:
        """
        Create a ``UnitAppFontWidth`` unit.

        Returns:
            UnitAppFontBase: The created unit.
        """
        from ooodev.units.unit_app_font_width import UnitAppFontWidth

        return UnitAppFontWidth(val)

    @staticmethod
    def create_app_font_x(val: float) -> UnitAppFontX:
        """
        Create a ``UnitAppFontX`` unit.

        Returns:
            UnitAppFontBase: The created unit.
        """
        from ooodev.units.unit_app_font_x import UnitAppFontX

        return UnitAppFontX(val)

    @staticmethod
    def create_app_font_y(val: float) -> UnitAppFontY:
        """
        Create a ``UnitAppFontY`` unit.

        Returns:
            UnitAppFontBase: The created unit.
        """
        from ooodev.units.unit_app_font_y import UnitAppFontY

        return UnitAppFontY(val)
