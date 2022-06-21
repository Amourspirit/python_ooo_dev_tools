# coding: utf-8
"""
Named Events.
"""
from typing import NamedTuple

class LoNamedEvent(NamedTuple):
    """
    Named events for utils.lo.LO class
    """
    DISPATCHING = "lo_dispatching"
    DISPATCHED = "lo_dispatched"
    DOC_CLOSING = "lo_doc_closing"
    DOC_CLOSED = "lo_doc_closed"
    DOC_CREATING = "lo_doc_creating"
    DOC_CREATED = "lo_doc_created"
    DOC_OPENING = "lo_doc_opening"
    DOC_OPENED = "lo_doc_opened"
    DOC_SAVING = "lo_doc_saving"
    DOC_SAVED = "lo_doc_saved"
    OFFICE_LOADING = "lo_office_loading"
    OFFICE_LOADED = "lo_office_loaded"
    OFFICE_CLOSING = "lo_office_closing"
    OFFICE_CLOSED = "lo_office_closed"
    RESET = "lo_reset"