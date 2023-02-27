from __future__ import annotations
from typing import TypeVar, Type
from dataclasses import dataclass
from ..validation import check
from ..decorator import enforce
from .base_int_value import BaseIntValue
from ..unit_convert import UnitConvert, Length

_TUnitMM100 = TypeVar(name="_TUnitMM100", bound="UnitMM100")

# Note that from __future__ import annotations converts annotations to string.
# this means that @enforce.enforce_types will see string as type. This is fine in
# most cases. Especially for built in types.
@enforce.enforce_types
@dataclass(unsafe_hash=True)
class UnitMM100(BaseIntValue):
    """
    Represents ``1/100th mm`` units.

    Supports ``UnitObj`` protcol.

    See Also:
        :ref:`proto_unit_obj`
    """

    def __post_init__(self) -> None:
        check(
            self.value >= 0,
            f"{self}",
            f"Value of {self.value} is out of range. Value must be a positive number.",
        )

    def _from_int(self, value: int) -> _TUnitMM100:
        inst = super(UnitMM100, self.__class__).__new__(self.__class__)
        return inst.__init__(value)

    def __eq__(self, other: object) -> bool:
        # for some reason BaseIntValue __eq__ is not picked up.
        # I suspect this is due to this class being a dataclass.
        try:
            i = int(other)
            return i == self.value
        except Exception as e:
            return False

    def get_value_mm(self) -> float:
        """
        Gets instance value converted to ``mm`` units.

        Returns:
            int: Value in ``mm`` units.
        """
        return UnitConvert.convert_mm100_mm(self.value)

    def get_value_pt(self) -> float:
        """
        Gets instance value converted to ``pt`` (point) units.

        Returns:
            int: Value in ``pt`` units.
        """
        return UnitConvert.convert(num=self.value, frm=Length.MM100, to=Length.PT)

    def get_value_mm100(self) -> int:
        """
        Gets instance value in ``1/100th mm`` units.

        Returns:
            int: Value in ``1/100th mm`` units.
        """
        return self.value

    @classmethod
    def from_mm(cls: Type[_TUnitMM100], value: float) -> _TUnitMM100:
        """
        Get instance from ``mm`` value.

        Args:
            value (int): ``mm`` value.

        Returns:
            UnitMM100:
        """
        inst = super(UnitMM100, cls).__new__(cls)
        return inst.__init__(UnitConvert.convert_mm_mm100(value))

    @classmethod
    def from_mm10(cls: Type[_TUnitMM100], value: float) -> _TUnitMM100:
        """
        Get instance from ``1/10th mm`` value.

        Args:
            value (int): ``1/10th mm`` value.

        Returns:
            UnitMM100:
        """
        inst = super(UnitMM100, cls).__new__(cls)
        return inst.__init__(round(UnitConvert.convert(num=value, frm=Length.MM10, to=Length.MM100)))

    @classmethod
    def from_mm100(cls: Type[_TUnitMM100], value: int) -> _TUnitMM100:
        """
        Get instance from ``1/100th mm`` value.

        Args:
            value (int): ``1/100th mm`` value.

        Returns:
            UnitMM100:
        """
        inst = super(UnitMM100, cls).__new__(cls)
        return inst.__init__(value)

    @classmethod
    def from_pt(cls: Type[_TUnitMM100], value: float) -> _TUnitMM100:
        """
        Get instance from ``pt`` (points) value.

        Args:
            value (int): ``pt`` value.

        Returns:
            UnitMM100:
        """
        inst = super(UnitMM100, cls).__new__(cls)
        return inst.__init__(round(UnitConvert.convert(num=value, frm=Length.PT, to=Length.MM100)))

    @classmethod
    def from_inch(cls: Type[_TUnitMM100], value: float) -> _TUnitMM100:
        """
        Get instance from ``in`` (inch) value.

        Args:
            value (int): ``in`` value.

        Returns:
            UnitMM100:
        """
        inst = super(UnitMM100, cls).__new__(cls)
        return inst.__init__(round(UnitConvert.convert(num=value, frm=Length.IN, to=Length.MM100)))

    @classmethod
    def from_inch10(cls: Type[_TUnitMM100], value: float) -> _TUnitMM100:
        """
        Get instance from ``1/10th in`` (inch) value.

        Args:
            value (int): ```/10th in`` value.

        Returns:
            UnitMM100:
        """
        inst = super(UnitMM100, cls).__new__(cls)
        return inst.__init__(round(UnitConvert.convert(num=value, frm=Length.IN10, to=Length.MM100)))

    @classmethod
    def from_inch100(cls: Type[_TUnitMM100], value: float) -> _TUnitMM100:
        """
        Get instance from ``1/100th in`` (inch) value.

        Args:
            value (int): ``1/100th in`` value.

        Returns:
            UnitMM100:
        """
        inst = super(UnitMM100, cls).__new__(cls)
        return inst.__init__(round(UnitConvert.convert(num=value, frm=Length.IN100, to=Length.MM100)))

    @classmethod
    def from_inch1000(cls: Type[_TUnitMM100], value: int) -> _TUnitMM100:
        """
        Get instance from ``1/1,000th in`` (inch) value.

        Args:
            value (int): ``1/1,000th in`` value.

        Returns:
            UnitMM100:
        """
        inst = super(UnitMM100, cls).__new__(cls)
        return inst.__init__(round(UnitConvert.convert(num=value, frm=Length.IN1000, to=Length.MM100)))