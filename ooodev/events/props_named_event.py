# coding: utf-8
"""
Lo Named Events.
"""
from __future__ import annotations
from typing import NamedTuple


class PropsNamedEvent(NamedTuple):
    """
    Named events for utils.lo.LO class
    """

    PROP_SETTING = "props_setting"
    """Prop setting command see :py:meth:`Props.set() <.utils.props.Props.set>`"""
    PROP_SET = "props_set"
    """Prop set command see :py:meth:`Props.set() <.utils.props.Props.set>`"""

    PROP_DEFAULT_SETTING = "props_default_setting"
    """Prop setting default command see :py:meth:`Props.set_default() <.utils.props.Props.set_default>`"""
    PROP_DEFAULT_SET = "props_default_set"
    """Prop set default command see :py:meth:`Props.set_default() <.utils.props.Props.set_default>`"""
