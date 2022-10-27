class GlobalEditDispatch:
    """
    Global Edit Dispatch Commands

    See Also:
        - :py:meth:`.Lo.dispatch_cmd`
        - `Global Dispatch commands <https://wiki.documentfoundation.org/Development/DispatchCommands#Global>`_
    """

    COMPARE_DOCUMENTS = "CompareDocuments"
    """
    Compare Non-Track Changed Document
    
    Args
        - ``URL (str)``
        - ``FilterName (str)``
        - ``Password (str)``
        - ``FilterOptions (str)``
        - ``Version (int)``
        - ``NoAcceptDialog (bool)``
    """
    COPY = "Copy"
    COPY_HYPERLINK_LOCATION = "CopyHyperlinkLocation"
    """Copy Hyperlink Location"""
    CUT = "Cut"
    DS_BROWSER_EXPLORER = "DSBrowserExplorer"
    """Explorer On/Off"""
    DELETE = "Delete"
    """Delete Contents"""
    DELETE_ALL_NOTES = "DeleteAllNotes"
    """Delete All Comments"""
    DELETE_AUTHOR = "DeleteAuthor"
    """Delete All Comments by This Author"""
    DELETE_COMMENT = "DeleteComment"
    """
    Delete Comment

    Args
        - ``ID (postitid)``
    """
    DELETE_COMMENT_THREAD = "DeleteCommentThread"
    """
    Delete Comment Thread

    Args
        - ``ID (postitid)``
    """
    DELETE_FRAME = "DeleteFrame"
    """Delete Frame"""
    EDIT_HYPERLINK = "EditHyperlink"
    """Edit Hyperlink"""
    EDIT_STYLE = "EditStyle"
    """
    Edit

    Args
        - ``Param (str)``
        - ``Family (int)``
    """
    FORMAT_ALL_NOTES = "FormatAllNotes"
    """Format All Comments"""
    GET_COLOR_TABLE = "GetColorTable"
    """Color Palette"""
    HYPERLINK_DIALOG = "HyperlinkDialog"
    """Insert Hyperlink"""
    INSERT_MODE = "InsertMode"
    """Insert Mode"""
    LEAVE_FM_CREATE_MODE = "LeaveFMCreateMode"
    MERGE_DOCUMENTS = "MergeDocuments"
    """
    Merge Track Changed Document

    Args
        - ``URL (str)``
        - ``Version (int)``
    """
    OBJECT_MENUE = "ObjectMenue"
    """Menu for editing or saving OLE objects"""
    OPEN_HYPERLINK_ON_CURSOR = "OpenHyperlinkOnCursor"
    """Open Hyperlink"""
    OPEN_SMART_TAG_MENU_ON_CURSOR = "OpenSmartTagMenuOnCursor"
    """Smart Tags"""
    PASTE = "Paste"
    PASTE_UNFORMATTED = "PasteUnformatted"
    """``Un-formatted`` Text"""
    PROTECT_TRACE_CHANGE_MODE = "ProtectTraceChangeMode"
    """Protect Track Changes"""
    REDO = "Redo"
    REFRESH = "Refresh"
    REPEAT = "Repeat"
    REPLY_COMMENT = "ReplyComment"
    """Reply Comment"""
    RESOLVE_COMMENT = "ResolveComment"
    """Resolved"""
    RESOLVE_COMMENT_THREAD = "ResolveCommentThread"
    """Resolved Thread"""
    SEARCH_DIALOG = "SearchDialog"
    """Find and Replace"""
    SEARCH_FORMATTED_DISPLAY_STRING = "SearchFormattedDisplayString"
    """Search Formatted Display String"""
    SEARCH_LABEL = "SearchLabel"
    """[placeholder for message]"""
    SELECT = "Select"
    """Select All"""
    SELECT_ALL = "SelectAll"
    """Select All"""
    SELECT_OBJECT = "SelectObject"
    """Select"""
    SPELL_DIALOG = "SpellDialog"
    """Spelling"""
    SPELLING_AND_GRAMMAR_DIALOG = "SpellingAndGrammarDialog"
    """Check Spelling"""
    SPELLING_DIALOG = "SpellingDialog"
    """Check Spelling"""
    UNDO = "Undo"
    UNDO_ACTION = "UndoAction"
