# coding: utf-8
"""
Calc Named Events.
"""
from __future__ import annotations
from typing import NamedTuple


class CalcNamedEvent(NamedTuple):
    """
    Named events for office.calc.Calc class
    """
    DOC_OPENING = "calc_doc_opening"
    DOC_OPENED = "calc_doc_opened"
    DOC_SS = "calc_doc_ss"
    DOC_CREATING = "calc_doc_creating"
    DOC_CREATED = "calc_doc_created"
    SHEET_GETTING = "calc_sheet_getting"
    SHEET_GET = "calc_sheet_get"