from __future__ import annotations
from enum import Enum
from ooodev.utils.data_type.size import Size


class PaperFormatKind(Enum):
    A6 = (1, 10500, 14800)
    A5 = (2, 14800, 21000)
    A4 = (3, 21000, 29700)
    A3 = (4, 29700, 42000)
    B6_ISO = (5, 12500, 17600)
    B5_ISO = (6, 17600, 25000)
    B4_ISO = (7, 25000, 35300)
    LETTER = (8, 21590, 27940)
    LEGAL = (9, 21590, 35560)
    LONG_BOND = (10, 21590, 33020)
    TABLOID = (11, 27940, 43180)
    B6_JIS = (12, 12800, 18200)
    B5_JIS = (13, 18200, 25700)
    B4_JIS = (14, 25700, 36400)
    KAI_16 = (15, 18400, 26000)
    KAI_32 = (16, 13000, 18400)
    KAI_32_BIG = (17, 14000, 20300)
    ENVELOPE_DL = (18, 11000, 22000)
    ENVELOPE_C6 = (19, 11400, 16200)
    ENVELOPE_C6_C5 = (20, 11400, 22900)
    ENVELOPE_C5 = (21, 16200, 22900)
    ENVELOPE_C4 = (22, 22900, 32400)
    ENVELOPE_6_3X4 = (23, 9208, 16510)
    ENVELOPE_7_3X4 = (24, 9843, 19050)
    ENVELOPE_9 = (25, 9843, 22543)
    ENVELOPE_10 = (26, 10478, 24130)
    ENVELOPE_11 = (27, 11430, 26353)
    ENVELOPE_12 = (28, 12065, 27940)
    POSTCARD_JAPANESE = (29, 10000, 14800)

    def get_size(self) -> Size:
        return Size(self.value[1], self.value[2])

    def __int__(self) -> int:
        return self.value[0]
