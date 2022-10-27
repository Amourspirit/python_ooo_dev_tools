class GlobalFormatDispatch:
    """
    Global Format Dispatch Commands

    See Also:
        - :py:meth:`.Lo.dispatch_cmd`
        - `Global Dispatch commands <https://wiki.documentfoundation.org/Development/DispatchCommands#Global>`_
    """

    ALIGN_CENTER = "AlignCenter"
    """Centered"""
    ALIGN_DOWN = "AlignDown"
    """Bottom"""
    ALIGN_MIDDLE = "AlignMiddle"
    """Center"""
    ALIGN_UP = "AlignUp"
    """Top"""
    AUTO_FORMAT = "AutoFormat"
    """
    Auto-Format Styles

    Args
        - ``aFormatName (str)``
    """
    BACKGROUND_COLOR = "BackgroundColor"
    """
    Background Color

    Args
        - ``Color (str)``
        - ``BackgroundColor  (color)``
    """
    BACKGROUND_PATTERN_CONTROLLER = "BackgroundPatternController"
    """Background Pattern"""
    BOLD = "Bold"
    """Bold"""
    BOLD_LATIN = "BoldLatin"
    BRING_TO_FRONT = "BringToFront"
    """Bring to Front"""
    CENTER_PARA = "CenterPara"
    """Align Center"""
    CHANGE_CASE_ROTATE_CASE = "ChangeCaseRotateCase"
    """
    Cycle Case ``(Title Case, Sentence case, UPPERCASE, lowercase)``
    """
    CHANGE_CASE_TO_FULL_WIDTH = "ChangeCaseToFullWidth"
    """Full-width"""
    CHANGE_CASE_TO_HALF_WIDTH = "ChangeCaseToHalfWidth"
    """Half-width"""
    CHANGE_CASE_TO_HIRAGANA = "ChangeCaseToHiragana"
    """Hiragana"""
    CHANGE_CASE_TO_KATAKANA = "ChangeCaseToKatakana"
    """Katakana"""
    CHANGE_CASE_TO_LOWER = "ChangeCaseToLower"
    """``lowercase``"""
    CHANGE_CASE_TO_SENTENCE_CASE = "ChangeCaseToSentenceCase"
    """Sentence case"""
    CHANGE_CASE_TO_TITLE_CASE = "ChangeCaseToTitleCase"
    """Capitalize Every Word"""
    CHANGE_CASE_TO_TOGGLE_CASE = "ChangeCaseToToggleCase"
    """``tOGGLE cASE``"""
    CHANGE_CASE_TO_UPPER = "ChangeCaseToUpper"
    """``UPPERCASE``"""
    CHAR_BACK_COLOR = "CharBackColor"
    """
    Args
        - ``Color (str)``
        - ``CharBackColor (color)``
    """
    CHAR_FONT_NAME = "CHAR_FONT_NAME"
    """FONT NAME"""
    CHAR_FONT_NAME_LATIN = "CharFontNameLatin"
    CHARACTER_BACKGROUND_PATTERN = "CharacterBackgroundPattern"
    """Highlight Color"""
    COLOR = "Color"
    """
    Font Color

    Args
        - ``Color (str)``
        - ``Color (color)``
        - ``ColorThemeIndex (int)``
        - ``ColorLumMod (int)``
        - ``ColorLumOff (int)``
    """
    COMMON_ALIGN_BOTTOM = "CommonAlignBottom"
    """Bottom"""
    COMMON_ALIGN_HORIZONTAL_CENTER = "CommonAlignHorizontalCenter"
    """Centered"""
    COMMON_ALIGN_HORIZONTAL_DEFAULT = "CommonAlignHorizontalDefault"
    """Default"""
    COMMON_ALIGN_JUSTIFIED = "CommonAlignJustified"
    """Justified"""
    COMMON_ALIGN_LEFT = "CommonAlignLeft"
    """Left"""
    COMMON_ALIGN_RIGHT = "CommonAlignRight"
    """Right"""
    COMMON_ALIGN_TOP = "CommonAlignTop"
    """Top"""
    COMMON_ALIGN_VERTICAL_CENTER = "CommonAlignVerticalCenter"
    """Center"""
    COMMON_ALIGN_VERTICAL_DEFAULT = "CommonAlignVerticalDefault"
    COMMON_TASK_BAR_VISIBLE = "CommonTaskBarVisible"
    """Presentation"""
    """Default"""
    DECREMENT_INDENT = "DecrementIndent"
    """Decrease Indent"""
    EMPHASIS_MARK = "EmphasisMark"
    ENTER_GROUP = "EnterGroup"
    """Enter Group"""
    FILL_COLOR = "FillColor"
    """
    Fill Color

    Args
        - ``Color (str)``
        - ``FillColor (xfillcolor)``
        - ``ColorThemeIndex (int)``
        - ``ColorLumMod (int)``
        - ``ColorLumOff (int)``
    """
    FILL_FLOAT_TRANSPARENCE = "FillFloatTransparence"
    """Gradient Fill Transparency"""
    FILL_STYLE = "FillStyle"
    """Area Style / Filling"""
    FILL_TRANSPARENCE = "FillTransparence"
    """Fill Transparency"""
    FILTER_CRIT = "FilterCrit"
    """Standard Filter"""
    FIND_ALL = "FindAll"
    """Find All"""
    FIND_TEXT = "FindText"
    """Find text in values, to search in formulas use the dialog"""
    FIT_CELL_SIZE = "FitCellSize"
    "Fit to Cell Size"
    FLIP_HORIZONTAL = "FlipHorizontal"
    """Flip Horizontally"""
    FLIP_VERTICAL = "Flip Vertically"
    FONT_HEIGH_LATIN = "FontHeighLatin"
    FONT_HEIGHT = "FontHeight"
    """Font Size"""
    FONT_WORK = "FontWork"
    FORMAT_AREA = "FormatArea"
    FORMAT_BULLETS_MENU = "FormatBulletsMenu"
    FORMAT_FORM_MENU = "FormatFormMenu"
    FORMAT_FRAME_MENU = "FormatFrameMenu"
    FORMAT_GROUP = "FormatGroup"
    """Group"""
    FORMAT_IMAGE_FILTERS_MENU = "FormatImageFiltersMenu"
    """Filter"""
    FORMAT_IMAGE_MENU = "FormatImageMenu"
    """Image"""
    FORMAT_LINE = "FormatLine"
    """Line"""
    FORMAT_SPACING_MENU = "FormatSpacingMenu"
    """Spacing"""
    FORMAT_STYLES_MENU = "FormatStylesMenu"
    """Styles"""
    FORMAT_TEXT_MENU = "FormatTextMenu"
    """Text"""
    FORMAT_UNGROUP = "FormatUngroup"
    """``Un-group``"""
    GRAF_BLUE = "GrafBlue"
    """Blue"""
    GRAF_CONTRAST = "GrafContrast"
    """Contrast"""
    GRAF_GAMMA = "GrafGamma"
    """Gamma"""
    GRAF_GREEN = "GrafGreen"
    """Green"""
    GRAF_INVERT = "GrafInvert"
    """Invert"""
    GRAF_LUMINANCE = "GrafLuminance"
    """Brightness"""
    GRAF_MODE = "GrafMode"
    """Image Mode"""
    GRAF_RED = "GrafRed"
    """Red"""
    GRAF_TRANSPARENCE = "GrafTransparence"
    """Transparency"""
    ITALIC = "Italic"
    ITALIC_LATIN = "ItalicLatin"
    """Italic Latin"""
    JUSTIFY_PARA = "JustifyPara"
    """Justified"""
    LANGUAGE_LATIN = "LanguageLatin"
    LANGUAGE_STATUS = "LanguageStatus"
    """Language Status"""
    LEAVEGROUP = "LeaveGroup"
    """Exit Group"""
    LEFT_PARA = "LeftPara"
    """Align Left"""
    LINE_CAP = "LineCap"
    LINE_DASH = "LineDash"
    LINE_END_STYLE = "LineEndStyle"
    """
    Select start and end arrowheads for lines.

    Args
        - ``LineStart (xlinestart)``
        - ``LineEnd (xlineend)``
        - ``StartWidth (int)``
        - ``EndWidth (int)``
    """
    LINE_JOINT = "LineJoint"
    """Line Corner Style"""
    LINE_SPACING = "LineSpacing"
    """Set Line Spacing"""
    LINE_TRANSPARENCE = "LineTransparence"
    """Line Transparency"""
    LINE_WIDTH = "LineWidth"
    """
    Line Width

    Args
        - ``Width (float)``
        - ``LineWidth (LineWidth)``
    """
    MEASURE_ATTRIBUTES = "MeasureAttributes"
    """Dimensions"""
    OUTLINE_FONT = "OutlineFont"
    """Apply outline attribute to font. Not all fonts implement this attribute."""
    OUTLINE_FORMAT = "OutlineFormat"
    """Show Formatting"""
    OVERLINE = "Overline"
    PARA_LEFT_TO_RIGHT = "ParaLeftToRight"
    """Left-To-Right"""
    PARA_RIGHT_TO_LEFT = "ParaRightToLeft"
    """Right-To-Left"""
    PARAGRAPH_DIALOG = "ParagraphDialog"
    """Paragraph"""
    PARASPACE_DECREASE = "ParaspaceDecrease"
    "Decrease Paragraph Spacing"
    PARASPACE_INCREASE = "ParaspaceIncrease"
    """Increase Paragraph Spacing"""
    RIGHT_PARA = "RightPara"
    """Align Right"""
    RUBY_DIALOG = "RubyDialog"
    """Asian Phonetic Guide"""
    RULER_MENU = "RulerMenu"
    """Rulers"""
    SEND_TO_BACK = "SendToBack"
    """Send to Back"""
    SET_DEFAULT = "SetDefault"
    """Clear Direct Formatting"""
    SET_OBJECT_TO_BACKGROUND = "SetObjectToBackground"
    """To Background"""
    SET_OBJECT_TO_FOREGROUND = "SetObjectToForeground"
    """To Foreground"""
    SHADOWED = "Shadowed"
    """Toggle Shadow"""
    SHAPES_LINE_MENU = "ShapesLineMenu"
    """Line"""
    SHAPES_MENU = "ShapesMenu"
    """Shape"""
    SHRINK = "Shrink"
    """Decrease Font Size"""
    SMALL_CAPS = "SmallCaps"
    """Small capitals"""
    SPACE_PARA_1 = "SpacePara1"
    """Line Spacing: 1"""
    SPACE_PARA_1_15 = "SpacePara115"
    """Line Spacing: 1.15"""
    SPACE_PARA_1_5 = "SpacePara15"
    """Line Spacing: 1.5"""
    SPACE_PARA_2 = "SpacePara2"
    "Line Spacing: 2"
    SPACING = "Spacing"
    """Set Character Spacing"""
    SPELL_CHECK_APPLY_SUGGESTION = "SpellCheckApplySuggestion"
    """
    Apply Suggestion

    Args
        - ``ApplyRule (str)``
    """
    SPELL_CHECK_IGNORE = "SpellCheckIgnore"
    """Ignore"""
    SPELL_CHECK_IGNORE_ALL = "SpellCheckIgnoreAll"
    """Ignore All"""
    STRIKE_OUT = "Strikeout"
    """Strike through"""
    TABLE_DESIGN = "TableDesign"
    """Table Design"""
    TABLE_STYLE_SETTINGS = "TableStyleSettings"
    """
    Args
        - ``UseFirstRowStyle (bool)``
        - ``UseLastRowStyle (bool)``
        - ``UseBandingRowStyle (bool)``
        - ``UseFirstColumnStyle (bool)``
        - ``UseLastColumnStyle (bool)``
        - ``UseBandingColumnStyle (bool)``
    """
    TEXT_FIT_TO_SIZE = "TextFitToSize"
    """Fit to Frame"""
    TEXT_DIRECTION_TOP_TO_BOTTOM = "TextdirectionTopToBottom"
    """Text direction from top to bottom"""
    TRANSFORM_DIALOG = "TransformDialog"
    """
    Position and Size

    Args
        - `TransformPosX (int)``
        - ``TransformPosY (int)``
        - ``TransformWidth (int)``
        - ``TransformHeight (int)``
        - ``TransformRotationDeltaAngle (sdrangle)``
        - ``TransformRotationAngle (sdrangle)``
        - `TransformRotationX (int)``
        - ``TransformRotationY (int)``
    """
    TRANSFORM_ROTATION_ANGLE = "TransformRotationAngle"
    """Rotation Angle"""
    TRANSFORM_ROTATION_X = "TransformRotationX"
    """Rotation Pivot Point X"""
    TRANSFORM_ROTATION_Y = "TransformRotationY"
    """Rotation Pivot Point Y"""
    UNDERLINE = "Underline"
    UNDERLINE_DOTTED = "UnderlineDotted"
    """Dotted Underline"""
    UNDERLINE_DOUBLE = "UnderlineDouble"
    """Double Underline"""
    UNDERLINE_NONE = "UnderlineNone"
    """Underline: Off"""
    UNDERLINE_SINGLE = "UnderlineSingle"
    """Single Underline"""
    X_LINE_COLOR = "XLineColor"
    """
    Line Color

    Args
        - `Color (str)``
        - ``XLineColor (xlinecolor)``
    """
    X_LINE_STYLE = "XLineStyle"
    """Line Style"""
