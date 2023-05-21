import math
from typing import List
from enum import IntFlag

# https://github.com/LibreOffice/core/blob/675a3e3d545c5811c35bbd189ff0a241a83443d0/svtools/source/control/ctrlbox.cxx


THINTHICK_SMALLGAP_LINE2 = 15.0
THINTHICK_SMALLGAP_GAP = 15.0
THINTHICK_LARGEGAP_LINE1 = 30.0
THINTHICK_LARGEGAP_LINE2 = 15.0
THICKTHIN_SMALLGAP_LINE1 = 15.0
THICKTHIN_SMALLGAP_GAP = 15.0
THICKTHIN_LARGEGAP_LINE1 = 15.0
THICKTHIN_LARGEGAP_LINE2 = 30.0
OUTSET_LINE1 = 15.0
INSET_LINE2 = 15.0


class BorderWidthImplFlags(IntFlag):
    CHANGE_LINE1 = 1 << 0
    CHANGE_LINE2 = 1 << 1
    CHANGE_DIST = 1 << 2


class BorderWidthImpl:
    """Border Width Implementation"""

    MINGAPWIDTH = 2

    def __init__(self, nFlags: BorderWidthImplFlags, nRate1: float, nRate2: float, nRateGap: float) -> None:
        self.m_n_flags = nFlags
        self.m_n_rate1 = nRate1
        self.m_n_rate2 = nRate2
        self.m_n_rate_gap = nRateGap

    def lcl_guessed_width(self, nTested: int, nRate: float, bChanging: bool) -> float:
        width = -1.0
        if bChanging:
            width = nTested / nRate
        elif math.isclose(float(nTested), nRate):
            width = nRate
        return width

    def guess_width(self, nLine1: int, nLine2: int, nGap: int) -> int:
        compare_lst: List[float] = []
        invalid = False
        line1_change = BorderWidthImplFlags.CHANGE_LINE1 in self.m_n_flags
        width1 = self.lcl_guessed_width(nLine1, self.m_n_rate1, line1_change)
        if line1_change:
            compare_lst.append(width1)
        else:
            invalid = True

        line2_change = BorderWidthImplFlags.CHANGE_LINE2 in self.m_n_flags
        width2 = self.lcl_guessed_width(nLine2, self.m_n_rate2, line2_change)
        if line1_change:
            compare_lst.append(width2)
        else:
            invalid = True

        gap_change = BorderWidthImplFlags.CHANGE_DIST in self.m_n_flags
        width_gap = self.lcl_guessed_width(nGap, self.m_n_rate_gap, gap_change)
        if gap_change and nGap >= BorderWidthImpl.MINGAPWIDTH:
            compare_lst.append(width_gap)
        elif not gap_change and width_gap < 0:
            invalid = True

        width = 0.0
        if not invalid and compare_lst:
            width = compare_lst[0]
            for elem in compare_lst[1:]:
                invalid = width != elem
                if invalid:
                    break
            width = 0.0 if invalid else nLine1 + nLine2 + nGap

        return round(width)

    def get_gap(self, nWidth: int) -> int:
        result = int(self.m_n_rate_gap)
        if BorderWidthImplFlags.CHANGE_DIST in self.m_n_flags:
            constant1 = 0 if BorderWidthImplFlags.CHANGE_LINE1 in self.m_n_flags else self.m_n_rate1
            constant2 = 0 if BorderWidthImplFlags.CHANGE_LINE2 in self.m_n_flags else self.m_n_rate2
            result = max(0, int(((self.m_n_rate_gap * nWidth) + 0.5) - (constant1 + constant2)))
        # Avoid having too small distances (less than 0.1pt)
        if result < BorderWidthImpl.MINGAPWIDTH and self.m_n_rate1 > 0 and self.m_n_rate2 > 0:
            result = BorderWidthImpl.MINGAPWIDTH

        return result

    def get_line1(self, nWidth: int) -> int:
        result = int(self.m_n_rate1)
        if BorderWidthImplFlags.CHANGE_LINE1 in self.m_n_flags:
            constant1 = 0 if BorderWidthImplFlags.CHANGE_LINE2 in self.m_n_flags else self.m_n_rate2
            constant_d = 0 if BorderWidthImplFlags.CHANGE_DIST in self.m_n_flags else self.m_n_rate_gap
            result = max(0, int(((self.m_n_rate1 * nWidth) + 0.5) - (constant1 + constant_d)))
        if result == 0 and self.m_n_rate1 > 0.0 and nWidth > 0:
            # hack to essentially treat 1 twip DOUBLE border

            result = 1  # as 1 twip SINGLE border

        return result

    def get_line2(self, nWidth: int) -> int:
        result = int(self.m_n_rate2)
        if BorderWidthImplFlags.CHANGE_LINE2 in self.m_n_flags:
            constant1 = 0 if BorderWidthImplFlags.CHANGE_LINE1 in self.m_n_flags else self.m_n_rate1
            constant_d = 0 if BorderWidthImplFlags.CHANGE_DIST in self.m_n_flags else self.m_n_rate_gap
            result = max(0, int(((self.m_n_rate2 * nWidth) + 0.5) - (constant1 + constant_d)))

        return result
