# coding: utf-8
"""
Write Named Events.
"""
from __future__ import annotations


class WriteNamedEvent:
    """
    Named events for :py:class:`~.office.wite.Write` class
    """

    DOC_OPENING = "write_doc_opening"
    """Doc Opening Write document see :py:meth:`Write.open_doc() <.office.write.Write.open_doc>`"""
    DOC_OPENED = "write_doc_opened"
    """Doc Opened Write document see :py:meth:`Write.open_doc() <.office.write.Write.open_doc>`"""

    DOC_CREATING = "write_doc_creating"
    """Doc Creating Write document see :py:meth:`Write.create_doc() <.office.write.Write.create_doc>`"""
    DOC_CREATED = "write_doc_created"
    """Doc Created Write document see :py:meth:`Write.create_doc() <.office.write.Write.create_doc>`"""

    DOC_CLOSING = "write_doc_closing"
    """Doc Closing Write document see :py:meth:`Write.close_doc() <.office.write.Write.close_doc>`"""
    DOC_CLOSED = "write_doc_closed"
    """Doc Closed Write document see :py:meth:`Write.close_doc() <.office.write.Write.close_doc>`"""

    DOC_SAVING = "write_doc_saving"
    """Doc Closing Write document see :py:meth:`Write.save_doc() <.office.write.Write.save_doc>`"""
    DOC_SAVED = "write_doc_saved"
    """Doc Closed Write document see :py:meth:`Write.save_doc() <.office.write.Write.save_doc>`"""

    DOC_TMPL_CREATING = "write_doc_tmpl_creating"
    """Doc Creating Write document see :py:meth:`Write.create_doc_from_template() <.office.write.Write.create_doc_from_template>`"""
    DOC_TMPL_CREATED = "write_doc_tmpl_created"
    """Doc Created Write document see :py:meth:`Write.create_doc_from_template() <.office.write.Write.create_doc_from_template>`"""

    WORD_SELECTING = "write_word_selecting"
    """Word Selecting see :py:meth:`Write.select_next_word() <.office.write.Write.select_next_word>`"""
    WORD_SELECTED = "write_word_selected"
    """Word Selected see :py:meth:`Write.select_next_word() <.office.write.Write.select_next_word>`"""

    DOC_TEXT = "write_text_doc"
    """Got a Write document see :py:meth:`Write.get_text_doc() <.office.write.Write.get_text_doc>`"""

    PAGE_FORMAT_SETTING = "write_page_format_setting"
    """Page Format Setting see :py:meth:`Write.set_page_format() <.office.write.Write.set_page_format>`"""
    PAGE_FORMAT_SET = "write_page_format_setting"
    """Page Format Set see :py:meth:`Write.set_page_format() <.office.write.Write.set_page_format>`"""

    FORMULA_ADDING = "write_formula_adding"
    """Formula Adding see :py:meth:`Write.add_formula() <.office.write.Write.add_formula>`"""
    FORMULA_ADDED = "write_formula_added"
    """Formula Added see :py:meth:`Write.add_formula() <.office.write.Write.add_formula>`"""

    HYPER_LINK_ADDING = "write_hyper_link_adding"
    """Hyperlink Adding see :py:meth:`Write.add_hyperlink() <.office.write.Write.add_hyperlink>`"""
    HYPER_LINK_ADDED = "write_hyper_link_added"
    """Hyperlink Added see :py:meth:`Write.add_hyperlink() <.office.write.Write.add_hyperlink>`"""

    BOOKMARK_ADDING = "write_bookmark_adding"
    """Bookmark Adding see :py:meth:`Write.add_bookmark() <.office.write.Write.add_bookmark>`"""
    BOOKMARK_ADDED = "write_bookmark_added"
    """Bookmark Added see :py:meth:`Write.add_bookmark() <.office.write.Write.add_bookmark>`"""

    TEXT_FRAME_ADDING = "write_text_frame_adding"
    """Text frame Adding see :py:meth:`Write.add_text_frame() <.office.write.Write.add_text_frame>`"""
    TEXT_FRAME_ADDED = "write_text_frame_added"
    """Text frame Added see :py:meth:`Write.add_text_frame() <.office.write.Write.add_text_frame>`"""

    TABLE_ADDING = "write_table_adding"
    """Table Adding see :py:meth:`Write.add_table() <.office.write.Write.add_table>`"""
    TABLE_ADDED = "write_table_added"
    """Table Added see :py:meth:`Write.add_table() <.office.write.Write.add_table>`"""

    IMAGE_LINK_ADDING = "write_image_link_adding"
    """Image Link Adding see :py:meth:`Write.add_image_link() <.office.write.Write.add_image_link>`"""
    IMAGE_LINK_ADDED = "write_image_link_added"
    """Image Link Added see :py:meth:`Write.add_image_link() <.office.write.Write.add_image_link>`"""

    IMAGE_SHAPE_ADDING = "write_image_shape_adding"
    """Image Shape Adding see :py:meth:`Write.add_image_shape() <.office.write.Write.add_image_shape>`"""
    IMAGE_SHAPE_ADDED = "write_image_shape_added"
    """Image Shape Added see :py:meth:`Write.add_image_shape() <.office.write.Write.add_image_shape>`"""

    CONFIGURED_SERVICES_SETTING = "write_configure_services_setting"
    """Services configuration setting see :py:meth:`Write.set_configured_services() <.office.write.Write.set_configured_services>`"""
    CONFIGURED_SERVICES_SET = "write_configure_services_set"
    """Services configuration set see :py:meth:`Write.set_configured_services() <.office.write.Write.set_configured_services>`"""

    STYLE_PREV_PARA_STYLES_SETTING = "write_prev_para_prop_setting"
    STYLE_PREV_PARA_STYLES_SET = "write_prev_para_prop_set"

    STYLING = "write_styling"
    """Styling"""
    STYLED = "write_styled"
    """Styled"""

    STYLE_PREV_PARA_PROP_SETTING = "write_prev_para_prop_setting"
    STYLE_PREV_PARA_PROP_SET = "write_prev_para_prop_set"

    EXPORTING_PAGE_PNG = "write_exporting_page_png"
    """
    Exporting a Write Page to image format of PNG.
    
    .. seealso::
    
        - :py:meth:`WriteTextViewCursor.export_page_png() <ooodev.write.WriteTextViewCursor.export_page_png>`
        - :py:class:`PagePng <ooodev.write.export.page_png.PagePng>`
    """
    EXPORTED_PAGE_PNG = "write_exported_page_png"
    """
    Exported a Write Page to image format of PNG.
    
    .. seealso::
    
        - :py:meth:`WriteTextViewCursor.export_page_png() <ooodev.write.WriteTextViewCursor.export_page_png>`
        - :py:class:`PagePng <ooodev.write.export.page_png.PagePng>`
    """

    EXPORTING_PAGE_JPG = "write_exporting_page_jpg"
    """
    Exporting a Write Page to image format of JPG.
    
    .. seealso::
    
        - :py:meth:`WriteTextViewCursor.export_page_jpg() <ooodev.write.WriteTextViewCursor.export_page_jpg>`
        - :py:class:`PageJpg <ooodev.write.export.page_jpg.PageJpg>`
    """
    EXPORTED_PAGE_JPG = "write_exported_page_jpg"
    """
    Exported a Write Page to image format of JPG.
    
    .. seealso::
    
        - :py:meth:`WriteTextViewCursor.export_page_jpg() <ooodev.write.WriteTextViewCursor.export_page_jpg>`
        - :py:class:`PageJpg <ooodev.write.export.page_jpg.PageJpg>`
    """

    CHARACTER_STYLE_APPLYING = "write_character_style_applying"
    CHARACTER_STYLE_APPLIED = "write_character_style_applied"
    TABLE_STYLE_APPLYING = "write_table_style_applying"
    TABLE_STYLE_APPLIED = "write_table_style_applied"
