# coding: utf-8
"""
Chart2 Named Events.
"""
from __future__ import annotations


class Chart2NamedEvent:
    """
    Named events for ``Chart2``

    .. versionadded:: 0.9.4
    """

    CONTROLLERS_LOCKING = "chart2_controllers_locking"
    """Controllers Locking see :py:meth:`Chart2.lock_controllers() <.office.chart2.Chart2.lock_controllers>`"""
    CONTROLLERS_LOCKED = "chart2_controllers_locked"
    """Controllers Locked see :py:meth:`Chart2.lock_controllers() <.office.chart2.Chart2.lock_controllers>`"""

    CONTROLLERS_UNLOCKING = "chart2_controllers_unlocking"
    """Controllers UnLocking see :py:meth:`Chart2.unlock_controllers() <.office.chart2.Chart2.unlock_controllers>`"""
    CONTROLLERS_UNLOCKED = "chart2_controllers_unlocked"
    """Controllers UnLocked see :py:meth:`Chart2.unlock_controllers() <.office.chart2.Chart2.unlock_controllers>`"""
