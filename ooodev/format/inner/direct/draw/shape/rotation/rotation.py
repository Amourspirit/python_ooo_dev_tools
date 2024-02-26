from __future__ import annotations
from typing import Any, cast, Tuple, TYPE_CHECKING, TypeVar, Type, overload
import uno

from ooodev.exceptions import ex as mEx
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.units.angle100 import Angle100
from ooodev.utils import props as mProps

if TYPE_CHECKING:
    from ooodev.units.angle_unit_obj import AngleUnitT

_TRotation = TypeVar(name="_TRotation", bound="Rotation")

# pivot point and base point are difficult to calculate.
# pivot point is the point that the shape rotates around.
# base point is the point that the shape is positioned on the page.
# the Shape position change when the shape is rotated.
# this means shape.getPosition() will show a different value then the dialog box after a rotation.
# Maybe a shape's rotation should be set to zero before setting the position. Not sure about this.
# Because the values in the dialog box seem to be calculated on the fly will not implement the same here for now.


class Rotation(StyleBase):
    """
    Rotation of a shape.

    .. versionadded:: 0.17.4
    """

    def __init__(self, rotation: int | AngleUnitT = 0) -> None:
        """
        Constructor

        Args:
            rotation (int, AngleUnitT, optional): Specifies the rotation angle of the shape in degrees.
                Default is ``0``.

        Returns:
            None:
        """
        super().__init__()

        self.prop_rotation = rotation

    # region override
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.drawing.Shape",)
        return self._supported_services_values

    # endregion override
    # region from_obj()
    @overload
    @classmethod
    def from_obj(cls: Type[_TRotation], obj: object) -> _TRotation: ...

    @overload
    @classmethod
    def from_obj(cls: Type[_TRotation], obj: object, **kwargs) -> _TRotation: ...

    @classmethod
    def from_obj(cls: Type[_TRotation], obj: Any, **kwargs) -> _TRotation:
        """
        Creates a new instance from ``obj``.

        Args:
            obj (Any): UNO Shape object.

        Returns:
            Rotation: New instance.
        """
        inst = cls(**kwargs)

        if not inst._is_valid_obj(obj):
            raise mEx.NotSupportedError("Object is not supported for conversion to Position")

        angle = cast(int, mProps.Props.get(obj, "RotateAngle", 0))
        inst.prop_rotation = Angle100(angle)
        return inst

    # endregion from_obj()

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.SHAPE
        return self._format_kind_prop

    @property
    def prop_rotation(self) -> Angle100:
        """
        Gets/Sets Rotation angle of the shape in degrees in ``1/100 units``.

        Property can be set by passing int in degrees or AngleUnitT object.
        """
        # in 1/10 degree units
        pv = cast(int, self._get("RotateAngle"))
        return Angle100(pv)

    @prop_rotation.setter
    def prop_rotation(self, value: int | AngleUnitT) -> None:
        # in 1/10 degree units
        try:
            val = value.get_angle100()  # type: ignore
        except AttributeError:
            val = Angle100.from_angle(value).value  # type: ignore
        self._set("RotateAngle", val)

    # endregion Properties
