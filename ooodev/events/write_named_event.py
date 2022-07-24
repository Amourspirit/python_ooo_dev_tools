# coding: utf-8
"""
Calc Named Events.
"""
from __future__ import annotations
from typing import NamedTuple


class WriteNamedEvent(NamedTuple):
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

    PAGE_FORRMAT_SETTING = "write_page_format_setting"
    """Page Format Setting see :py:meth:`Write.set_page_format() <.office.write.Write.set_page_format>`"""
    PAGE_FORRMAT_SET = "write_page_format_setting"
    """Page Format Set see :py:meth:`Write.set_page_format() <.office.write.Write.set_page_format>`"""

    FORMULA_ADDING = "write_forumla_adding"
    """Formula Adding see :py:meth:`Write.add_formula() <.office.write.Write.add_formula>`"""
    FORMULA_ADDED = "write_forumla_added"
    """Formula Added see :py:meth:`Write.add_formula() <.office.write.Write.add_formula>`"""

    HYPER_LINK_ADDING = "write_hyper_link_adding"
    """Hyperlink Adding see :py:meth:`Write.add_hyperlink() <.office.write.Write.add_hyperlink>`"""
    HYPER_LINK_ADDED = "write_hyper_link_added"
    """Hyperlink Added see :py:meth:`Write.add_hyperlink() <.office.write.Write.add_hyperlink>`"""

    BOOKMARK_ADDING = "write_bookmark_adding"
    """Bookmark Adding see :py:meth:`Write.add_bookmark() <.office.write.Write.add_bookmark>`"""
    BOOKMARK_ADDIED = "write_bookmark_added"
    """Bookmark Added see :py:meth:`Write.add_bookmark() <.office.write.Write.add_bookmark>`"""

    TEXT_FRAME_ADDING = "write_text_frame_adding"
    """Text frame Adding see :py:meth:`Write.add_text_frame() <.office.write.Write.add_text_frame>`"""
    TEXT_FRAME_ADDED = "write_text_frame_added"
    """Text frame Added see :py:meth:`Write.add_text_frame() <.office.write.Write.add_text_frame>`"""

    TABLE_ADDING = "write_tabel_adding"
    """Table Adding see :py:meth:`Write.add_table() <.office.write.Write.add_table>`"""
    TABLE_ADDED = "write_tabel_added"
    """Table Added see :py:meth:`Write.add_table() <.office.write.Write.add_table>`"""

    IMAGE_LNIK_ADDING = "write_image_link_adding"
    """Image Link Adding see :py:meth:`Write.add_image_link() <.office.write.Write.add_image_link>`"""
    IMAGE_LNIK_ADDED = "write_image_link_added"
    """Image Link Added see :py:meth:`Write.add_image_link() <.office.write.Write.add_image_link>`"""

    IMAGE_SHAPE_ADDING = "write_image_shape_adding"
    """Image Shape Adding see :py:meth:`Write.add_image_shape() <.office.write.Write.add_image_shape>`"""
    IMAGE_SHAPE_ADDED = "write_image_shape_added"
    """Image Shape Added see :py:meth:`Write.add_image_shape() <.office.write.Write.add_image_shape>`"""

    CONFIGURED_SERVICES_SETTING = "write_configure_services_setting"
    """Services configuration setting see :py:meth:`Write.set_configured_services() <.office.write.Write.set_configured_services>`"""
    CONFIGURED_SERVICES_SET = "write_configure_services_set"
    """Services configuration set see :py:meth:`Write.set_configured_services() <.office.write.Write.set_configured_services>`"""
