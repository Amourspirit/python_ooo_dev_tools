from enum import Enum


class FamilyNamesKind(Enum):
    """Style Look ups for Write Families Styles"""

    CELL_STYLES = "CellStyles"
    CHARACTER_STYLES = "CharacterStyles"
    FRAME_STYLES = "FrameStyles"
    NUMBERING_STYLES = "NumberingStyles"
    PAGE_STYLES = "PageStyles"
    PARAGRAPH_STYLES = "ParagraphStyles"
    TABLE_STYLES = "TableStyles"

    def __str__(self) -> str:
        return self.value
