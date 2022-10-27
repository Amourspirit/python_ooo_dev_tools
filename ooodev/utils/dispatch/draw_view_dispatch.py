class DrawViewDispatch:
    """
    Draw View Dispatch Commands

    See Also:
        - :py:meth:`.Lo.dispatch_cmd`
        - `Draw Dispatch commands <https://wiki.documentfoundation.org/Development/DispatchCommands#Draw_.2F_Impress>`_
    """

    CLEAR_UNDO_STACK = "ClearUndoStack"
    COLOR_VIEW = "ColorView"
    """Black & White View"""
    DIA_MODE = "DiaMode"
    """Slide Sorter"""
    DISPLAY_MASTER_BACKGROUND = "DisplayMasterBackground"
    """Master Background"""
    GRID_FRONT = "GridFront"
    """Grid to Front"""
    HELPLINES_FRONT = "HelplinesFront"
    """Snap Guides to Front"""
    HELPLINES_USE = "HelplinesUse"
    """Snap to Snap Guides"""
    HELPLINES_VISIBLE = "HelplinesVisible"
    """Display Snap Guides"""
    LAYER_MODE = "LayerMode"
    """
    Layer
    
    Args
        - IsActive (bool)
        - WhatLayer (int)
    """
    LAYOUT_STATUS = "LayoutStatus"
    """Layout"""
    MASTER_PAGE = "MasterPage"
    """
    Master
    
    Args
        - IsActive (bool)
    """
    NEXT_ANNOTATION = "NextAnnotation"
    NOTES_CHILD_WINDOW = "NotesChildWindow"
    NOTES_MASTER_PAGE = "NotesMasterPage"
    """Master Notes"""
    NOTES_MODE = "NotesMode"
    """Notes"""
    OBJECT_POSITION = "ObjectPosition"
    """Arrange"""
    OUTLINE_MODE = "OutlineMode"
    """Outline"""
    OUTPUT_QUALITY_BLACK_WHITE = "OutputQualityBlackWhite"
    """Black and White"""
    OUTPUT_QUALITY_COLOR = "OutputQualityColor"
    """Color"""
    OUTPUT_QUALITY_CONTRAST = "OutputQualityContrast"
    """High Contrast"""
    OUTPUT_QUALITY_GRAYSCALE = "OutputQualityGrayscale"
    """Gray scale"""
    PAGE_MODE = "PageMode"
    """
    Normal
    
    Args
        - IsActive (bool)
        - WhatKind (int)
    """
    PAGE_STATUS = "PageStatus"
    """Slide/Layer"""
    PRESENTATION = "Presentation"
    """Start from First Slide"""
    PRESENTATION_CURRENT_SLIDE = "PresentationCurrentSlide"
    """Start from Current Slide"""
    PRESENTATION_END = "PresentationEnd"
    PREVIEW_STATE = "PreviewState"
    PREVIOUS_ANNOTATION = "PreviousAnnotation"
    SCALE = "Scale"
    SLIDE_MASTER_PAGE = "SlideMasterPage"
    """Master Slide"""
    SNAP_BORDER = "SnapBorder"
    """Snap to Page Margins"""
    SNAP_FRAME = "SnapFrame"
    SNAP_POINTS = "SnapPoints"
    """Snap to Object Points"""
    SOLID_CREATE = "SolidCreate"
    """Modify Object with Attributes"""
    SWITCH_LAYER = "SwitchLayer"
    """
    Args
        - WhatLayer (int)
    """
    SWITCH_PAGE = "SwitchPage"
    """
    Args
        - WhatPage (int)
        - WhatKind (int)
    """
    TOGGLE_TAB_BAR_VISIBILITY = "ToggleTabBarVisibility"
    """Toggle Views Tab Bar Visibility"""
    ZOOM_MODE = "ZoomMode"
    """Zoom & Pan (Control to Zoom Out, Shift to Pan)"""
    ZOOM_PANNING = "ZoomPanning"
    """Shift"""
