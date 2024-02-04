from __future__ import annotations
from dataclasses import dataclass


@dataclass(frozen=True)
class Options:
    """
    Lo Load options

    .. versionadded:: 0.6.10
    """

    verbose: bool = False
    """Determines if various info is sent to console. Default ``False``"""

    dynamic: bool = True
    """
    Determines if the script context is dynamic.

    Also When loading a component via :py:meth:`LoInst.load_component() <ooodev.utils.inst.lo.LoInst.load_component>` It is recommended to set this value to ``False``.

    If dynamic the script context is created with the current document; Otherwise, context is static.
    Static context is useful when the script is only going to operate on a a single document.
    There may be a performance gain when using static context.

    Default ``True``
    
    .. versionadded:: 0.11.13
    """
