"""
Module for converting between different units.

.. versionadded:: 0.9.0
"""

from __future__ import annotations
from typing import List, Tuple, NamedTuple, overload, Union
from enum import IntEnum
import math
from ooodev.utils import table_helper as mTh

# See Also:
#   https://github.com/LibreOffice/core/blob/e5005c76bd60a004f6025728e794ba3e4d0dfff1/include/o3tl/unit_conversion.hxx
#   https://help.libreoffice.org/latest/en-US/text/scalc/01/func_convert.html?&DbPAR=CALC&System=UNIX
#   https://wiki.documentfoundation.org/Documentation/Calc_Functions/CONVERT
#   https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1util_1_1MeasureUnit.html
#   https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XUnitConversion.html
#   https://ask.libreoffice.org/t/how-to-set-a-dialog-always-at-the-center-of-the-screen/97196/3
#   ooodev.adapter.awt.unit_conversion_partial.UnitConversionPartial
#   ooodev.dialog.dl_control.ctl_dialog.CtlDialog

N = Union[int, float]


class UnitLength(IntEnum):
    MM100 = 0
    """``1/100th mm``"""
    MM10 = 1
    """``1/10 mm``"""
    MM = 2
    """millimeter"""
    CM = 3
    """centimeter"""
    M = 4
    """meter"""
    KM = 5
    """kilometer"""
    EMU = 6
    """English Metric Unit: ``1/360000 cm``, ``1/914400 in``"""
    TWIP = 7
    """Twentieth of a point aka ``dxa``: ``1/20 pt``"""
    PT = 8
    """Point: ``1/72 in``"""
    PC = 9
    """Pica: ``1/6 in``"""
    IN1000 = 10
    """``1/1000 in``"""
    IN100 = 11
    """``1/100 in``"""
    IN10 = 12
    """``1/10 in``"""
    IN = 13
    """inch"""
    FT = 14
    """foot"""
    MI = 15
    """mile"""
    MASTER = 16
    """PPT Master Unit: ``1/576 in``"""
    PX = 17
    """pixel unit: ``15`` twip (``96 ppi``)"""
    CH = 18
    """char unit: ``210`` twip (``14 px``)"""
    LINE = 19
    """line unit: ``312`` twip"""
    COUNT = 20  # add new units above this last entry and add 1 for each new entry
    INVALID = -1
    """invalid"""


class MulDiv:
    def __init__(self, mul: int, div: int) -> None:
        # make sure to use smallest quotients here because
        # they will be multiplied when building final table
        self._mul = mul / UnitConvert.asserting_gcd(mul, div)
        self._div = div / UnitConvert.asserting_gcd(mul, div)

    @property
    def mul(self) -> float:
        """Gets mul value"""
        return self._mul

    @property
    def div(self) -> float:
        """Gets div value"""
        return self._div


class MdItem(NamedTuple):
    m: int
    d: int


_mul_div = (
    MdItem(1, 100),  # mm100 => mm
    MdItem(1, 10),  # mm10 => mm
    MdItem(1, 1),  # mm => mm
    MdItem(10, 1),  # cm => mm
    MdItem(1000, 1),  # m => mm
    MdItem(1000000, 1),  # km => mm
    MdItem(1, 36000),  # emu => mm
    MdItem(254, 10 * 1440),  # twip => mm
    MdItem(254, 10 * 72),  # pt => mm
    MdItem(254, 10 * 6),  # pc => mm
    MdItem(254, 10000),  # in1000 => mm
    MdItem(254, 1000),  # in100 => mm
    MdItem(254, 100),  # in10 => mm
    MdItem(254, 10),  # in => mm
    MdItem(254 * 12, 10),  # ft => mm
    MdItem(254 * 12 * 5280, 10),  # mi => mm
    MdItem(254, 10 * 576),  # master => mm
    MdItem(254 * 15, 10 * 1440),  # px => mm
    MdItem(254 * 210, 10 * 1440),  # ch => mm
    MdItem(254 * 312, 10 * 1440),  # line => mm
)


def _prepare_mul_div(md: Tuple[MdItem, ...]) -> List[List[int]]:
    n = len(md)
    a = mTh.TableHelper.make_2d_array(n, n)
    for i in range(n):
        a[i][i] = 1
        for j in range(i):
            m = md[i].m * md[j].d
            d = md[i].d * md[j].m
            g = UnitConvert.asserting_gcd(m, d)
            a[i][j] = m / g
            a[j][i] = d / g
    return a


class UnitConvert:
    @staticmethod
    def make_unsigned(num: N) -> N:
        """
        Gets unsigned number

        Args:
            num (N): Number

        Raises:
            AssertionError: If ``num`` is a negative number.

        Returns:
            N: Value of ``num`` if positive number.
        """
        # https://github.com/LibreOffice/core/blob/bdbb5d0389642c0d445b5779fe2a18fda3e4a4d4/include/o3tl/safeint.hxx
        assert num >= 0
        return num

    @classmethod
    def _md(cls, i: UnitLength | int, j: UnitLength | int) -> int:
        ni = int(i)
        nj = int(j)
        al = len(_a_length_md_array)
        assert ni >= 0 and cls.make_unsigned(ni) < al
        assert nj >= 0 and cls.make_unsigned(nj) < al
        return _a_length_md_array[ni][nj]

    @staticmethod
    def mul_div(num: N, mul: N, div: N) -> float:
        """
        Multiplies and divides.

        Args:
            num (N): number
            mul (N): multiplier
            div (N): divisor

        Returns:
            float: Converted Number

        Note:
            Formula ``num * (mul / div)``
        """
        assert mul > 0 and div > 0
        return num * (mul / div)

    @staticmethod
    def asserting_gcd(m: int, n: int) -> int:
        """
        Find the greatest common divisor of the two integers

        Args:
            m (int): The first integer to find the GCD for
            n (int): The second integer to find the GCD for

        Raises:
            AssertionError: If GCD result equals ``0``.

        Returns:
            int: A value, representing the greatest common divisor (GCD) for two integers
        """
        ret = math.gcd(m, n)
        assert ret != 0
        return ret

    @overload
    @classmethod
    def convert(cls, num: N, frm: UnitLength, to: UnitLength) -> float:
        """
        Converts a number from one unit to another unit.

        Args:
            num (N): Number to convert such as a ``float`` or ``int``.
            frm (Length): Current number kind.
            to (Length): Kind to convert to.

        Returns:
            float: Converted number
        """
        ...

    @overload
    @classmethod
    def convert(cls, num: N, frm: int, to: int) -> float:
        """
        Converts number by calling ``mul_div()``.

        Args:
            num (N): Number to convert such as a ``float`` or ``int``.
            frm (int): multiplier.
            to (int): divisor.

        Returns:
            float: Converted number
        """
        ...

    @classmethod
    def convert(cls, num: N, frm: int | UnitLength, to: int | UnitLength) -> float:
        """
        Converts a number from one unit to another unit.

        Args:
            num (N): Number to convert such as a ``float`` or ``int``.
            frm (int | Length): Current number kind.
            to (int | Length): Kind to convert to.

        Returns:
            float: Converted number

        Note:
            If ``Length`` is not used then :py:meth:`~.UnitConvert.mul_div` is called directly.
        """
        if isinstance(frm, UnitLength):
            return cls.mul_div(num, cls._md(frm, to), cls._md(to, frm))
        return cls.mul_div(num, frm, to)

    @classmethod
    def to_twips(cls, num: N, frm: UnitLength) -> float:
        """
        Converts number to twips

        Args:
            num (N): Number to convert
            frm (Length): The number kind to convert to twips

        Returns:
            float: Converted number
        """
        # Convert to twips - for convenience as we do this a lot
        return cls.convert(num=num, frm=frm, to=UnitLength.TWIP)

    @classmethod
    def convert_twip_mm100(cls, num: N) -> float:
        """
        Converts twips to ``1/100th mm``

        Args:
            num (N): Number to convert


        Returns:
            float: Converted number
        """
        return cls.convert(num=num, frm=UnitLength.TWIP, to=UnitLength.MM100)

    @classmethod
    def convert_pt_mm100(cls, num: N) -> int:
        """
        Converts points to ``1/100th mm``

        Args:
            num (N): Number to convert

        Returns:
            float: Converted number
        """
        return round(cls.convert(num=num, frm=UnitLength.PT, to=UnitLength.MM100))

    @classmethod
    def convert_mm100_pt(cls, num: N) -> float:
        """
        Converts ``1/100th mm`` to points

        Args:
            num (N): Number to convert

        Returns:
            float: Converted number
        """
        return cls.convert(num=num, frm=UnitLength.MM100, to=UnitLength.PT)

    @classmethod
    def convert_mm_mm100(cls, num: N) -> int:
        """
        Converts ``1/100th mm`` to ``mm``

        Args:
            num (N): Number to convert

        Returns:
            float: Converted number
        """
        return round(cls.convert(num=num, frm=UnitLength.MM, to=UnitLength.MM100))

    @classmethod
    def convert_mm100_mm(cls, num: N) -> float:
        """
        Converts ``mm`` to ``1/100th mm``

        Args:
            num (N): Number to convert

        Returns:
            float: Converted number
        """
        return cls.convert(num=num, frm=UnitLength.MM100, to=UnitLength.MM)


assert len(_mul_div) == UnitLength.COUNT

# The resulting multipliers and divisors array
_a_length_md_array = _prepare_mul_div(_mul_div)

__all__ = ("UnitConvert", "MdItem", "MulDiv", "UnitLength")
