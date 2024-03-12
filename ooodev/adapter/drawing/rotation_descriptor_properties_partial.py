from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from ooodev.units.angle100 import Angle100

if TYPE_CHECKING:
    from com.sun.star.drawing import RotationDescriptor
    from ooodev.units.angle_t import AngleT


class RotationDescriptorPropertiesPartial:
    """
    Service Class

    This abstract service specifies the general characteristics of an optional rotation and shearing for a Shape.

    This service is deprecated, instead please use the Transformation property of the service Shape.

    .. deprecated::

        Class is deprecated.

    See Also:
        `API RotationDescriptor <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1drawing_1_1RotationDescriptor.html>`_
    """

    def __init__(self, component: RotationDescriptor) -> None:
        """
        Constructor

        Args:
            component (RotationDescriptor): UNO Component that implements ``com.sun.star.drawing.RotationDescriptor`` interface.
        """
        self.__component = component

    # region RotationDescriptor
    @property
    def rotate_angle(self) -> Angle100:
        """
        This is the angle for rotation of this Shape.

        The shape is rotated counter-clockwise around the center of the bounding box.

        This property contains an error, the rotation angle is mathematically inverted when
        You take into account that the Y-Axis of the coordinate system is pointing down.
        Please use the Transformation property of the service Shape instead.

        When is setting this property, the value can be an int (in ``1/100th of a degree``) or an ``AngleT``.

        Returns:
            Angle100: The angle for rotation of this Shape.

        Hint:
            ``Angle100`` can be imported from ``ooodev.units``

        See Also:
            `Drawings and Presentations - Rotating and Shearing <https://wiki.documentfoundation.org/Documentation/DevGuide/Drawing_Documents_and_Presentation_Documents#Rotating_and_Shearing>`__
        """
        return Angle100(self.__component.RotateAngle)

    @rotate_angle.setter
    def rotate_angle(self, value: int | AngleT) -> None:
        val = Angle100.from_unit_val(value)
        self.__component.RotateAngle = val.value

    @property
    def shear_angle(self) -> Angle100:
        """
        This is the amount of shearing for this Shape.

        The shape is sheared counter-clockwise around the center of the bounding box

        When is setting this property, the value can be an int (in ``1/100th of a degree``) or an ``AngleT``.

        Returns:
            Angle100: The angle for rotation of this Shape.

        Hint:
            ``Angle100`` can be imported from ``ooodev.units``

        See Also:
            `Drawings and Presentations - Rotating and Shearing <https://wiki.documentfoundation.org/Documentation/DevGuide/Drawing_Documents_and_Presentation_Documents#Rotating_and_Shearing>`__
        """
        return Angle100(self.__component.ShearAngle)

    @shear_angle.setter
    def shear_angle(self, value: int | AngleT) -> None:
        val = Angle100.from_unit_val(value)
        self.__component.ShearAngle = val.value

    # endregion RotationDescriptor
