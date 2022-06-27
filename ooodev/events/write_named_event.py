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
    """Doc Opened Write document see :py:meth:`Write.open_doc() <.office.Write.Write.open_doc>`"""