# coding: utf-8
"""
Calc Named Events.
"""
from typing import NamedTuple


class CalcNamedEvent(NamedTuple):
    """
    Named events for office.calc.Calc class
    """
    DOC_OPENING = "calc_doc_opening"
    DOC_OPENED = "calc_doc_opened"