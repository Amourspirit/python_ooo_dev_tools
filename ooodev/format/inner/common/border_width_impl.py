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
        self.m_nFlags = nFlags
        self.m_nRate1 = nRate1
        self.m_nRate2 = nRate2
        self.m_nRateGap = nRateGap

    def lcl_guessed_width(self, nTested: int, nRate: float, bChanging: bool) -> float:
        nWidth = -1.0
        if bChanging:
            nWidth = nTested / nRate
        else:
            if math.isclose(float(nTested), nRate):
                nWidth = nRate
        return nWidth

    def guess_width(self, nLine1: int, nLine2: int, nGap: int) -> int:
        aToCompare: List[float] = []
        bInvalid = False
        bLine1Change = BorderWidthImplFlags.CHANGE_LINE1 in self.m_nFlags
        nWidth1 = self.lcl_guessed_width(nLine1, self.m_nRate1, bLine1Change)
        if bLine1Change:
            aToCompare.append(nWidth1)
        else:
            bInvalid = True

        bLine2Change = BorderWidthImplFlags.CHANGE_LINE2 in self.m_nFlags
        nWidth2 = self.lcl_guessed_width(nLine2, self.m_nRate2, bLine2Change)
        if bLine1Change:
            aToCompare.append(nWidth2)
        else:
            bInvalid = True

        bGapChange = BorderWidthImplFlags.CHANGE_DIST in self.m_nFlags
        nWidthGap = self.lcl_guessed_width(nGap, self.m_nRateGap, bGapChange)
        if bGapChange and nGap >= BorderWidthImpl.MINGAPWIDTH:
            aToCompare.append(nWidthGap)
        elif not bGapChange and nWidthGap < 0:
            bInvalid = True

        nWidth = 0.0
        if not bInvalid and len(aToCompare) > 0:
            nWidth = aToCompare[0]
            if len(aToCompare) > 1:
                for el in aToCompare[1]:
                    bInvalid = nWidth != el
                    if bInvalid:
                        break
        nWidth = 0.0 if bInvalid else nLine1 + nLine2 + nGap
        return nWidth

    def get_gap(self, nWidth: int) -> int:
        result = int(self.m_nRateGap)
        if BorderWidthImplFlags.CHANGE_DIST in self.m_nFlags:
            nConstant1 = 0 if BorderWidthImplFlags.CHANGE_LINE1 in self.m_nFlags else self.m_nRate1
            nConstant2 = 0 if BorderWidthImplFlags.CHANGE_LINE2 in self.m_nFlags else self.m_nRate2
            result = max(0, int(((self.m_nRateGap * nWidth) + 0.5) - (nConstant1 + nConstant2)))
        # Avoid having too small distances (less than 0.1pt)
        if result < BorderWidthImpl.MINGAPWIDTH and self.m_nRate1 > 0 and self.m_nRate2 > 0:
            result = BorderWidthImpl.MINGAPWIDTH

        return result

    def get_line1(self, nWidth: int) -> int:
        result = int(self.m_nRate1)
        if BorderWidthImplFlags.CHANGE_LINE1 in self.m_nFlags:
            nConstant1 = 0 if BorderWidthImplFlags.CHANGE_LINE2 in self.m_nFlags else self.m_nRate2
            nConstantD = 0 if BorderWidthImplFlags.CHANGE_DIST in self.m_nFlags else self.m_nRateGap
            result = max(0, int(((self.m_nRate1 * nWidth) + 0.5) - (nConstant1 + nConstantD)))
        if result == 0 and self.m_nRate1 > 0.0 and nWidth > 0:
            # hack to essentially treat 1 twip DOUBLE border

            result = 1  # as 1 twip SINGLE border

        return result

    def get_line2(self, nWidth: int) -> int:
        result = int(self.m_nRate2)
        if BorderWidthImplFlags.CHANGE_LINE2 in self.m_nFlags:
            nConstant1 = 0 if BorderWidthImplFlags.CHANGE_LINE1 in self.m_nFlags else self.m_nRate1
            nConstantD = 0 if BorderWidthImplFlags.CHANGE_DIST in self.m_nFlags else self.m_nRateGap
            result = max(0, int(((self.m_nRate2 * nWidth) + 0.5) - (nConstant1 + nConstantD)))

        return result
