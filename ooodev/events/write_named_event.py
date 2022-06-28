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