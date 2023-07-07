# coding: utf-8
"""
Lo Named Events.
"""
from __future__ import annotations


class LoNamedEvent:
    """
    Named events for utils.lo.LO class
    """

    DISPATCHING = "lo_dispatching"
    """Lo dispatching command see :py:meth:`Lo.dispatch_cmd() <.utils.lo.Lo.dispatch_cmd>`"""
    DISPATCHED = "lo_dispatched"
    """Lo dispatched command see :py:meth:`Lo.dispatch_cmd() <.utils.lo.Lo.dispatch_cmd>`"""

    DOC_CLOSING = "lo_doc_closing"
    """Lo Closing document see :py:meth:`Lo.close() <.utils.lo.Lo.close>`"""
    DOC_CLOSED = "lo_doc_closed"
    """Lo Closed document see :py:meth:`Lo.close() <.utils.lo.Lo.close>`"""

    DOC_CREATING = "lo_doc_creating"
    """Lo creating document see :py:meth:`Lo.create_doc() <.utils.lo.Lo.create_doc>`"""
    DOC_CREATED = "lo_doc_created"
    """Lo created document see :py:meth:`Lo.create_doc() <.utils.lo.Lo.create_doc>`"""

    DOC_OPENING = "lo_doc_opening"
    """Lo opening document see :py:meth:`Lo.open_doc() <.utils.lo.Lo.open_doc>`"""
    DOC_OPENED = "lo_doc_opened"
    """Lo opened document see :py:meth:`Lo.open_doc() <.utils.lo.Lo.open_doc>`"""

    DOC_SAVING = "lo_doc_saving"
    """
    Lo saving document
    see :py:meth:`Lo.save() <.utils.lo.Lo.save>`,
    :py:meth:`Lo.save_doc() <.utils.lo.Lo.save_doc>`,
    """
    DOC_SAVED = "lo_doc_saved"
    """
    Lo saved document
    see :py:meth:`Lo.save() <.utils.lo.Lo.save>`,
    :py:meth:`Lo.save_doc() <.utils.lo.Lo.save_doc>`,
    """

    DOC_STORING = "lo_doc_storing"
    """Lo Storing document see
    :py:meth:`Lo.store_doc() <.utils.lo.Lo.store_doc>`,
    :py:meth:`Lo.store_doc_format() <.utils.lo.Lo.store_doc_format>`"""
    DOC_STORED = "lo_doc_stored"
    """Lo Stored document see
    :py:meth:`Lo.store_doc() <.utils.lo.Lo.store_doc>`,
    :py:meth:`Lo.store_doc_format() <.utils.lo.Lo.store_doc_format>`
    """

    OFFICE_LOADING = "lo_office_loading"
    """Lo loading see :py:meth:`Lo.load_office() <.utils.lo.Lo.load_office>`"""
    OFFICE_LOADED = "lo_office_loaded"
    """Lo loaded see :py:meth:`Lo.load_office() <.utils.lo.Lo.load_office>`"""

    OFFICE_CLOSING = "lo_office_closing"
    """Lo closing office see :py:meth:`Lo.close_office() <.utils.lo.Lo.close_office>`"""
    OFFICE_CLOSED = "lo_office_closed"
    """Lo closed office see :py:meth:`Lo.close_office() <.utils.lo.Lo.close_office>`"""

    CONTROLLERS_LOCKING = "lo_controllers_locking"
    """Controllers Locking see :py:meth:`Lo.lock_controllers() <.utils.lo.Lo.lock_controllers>`"""
    CONTROLLERS_LOCKED = "lo_controllers_locked"
    """Controllers Locked see :py:meth:`Lo.lock_controllers() <.utils.lo.Lo.lock_controllers>`"""

    CONTROLLERS_UNLOCKING = "lo_controllers_unlocking"
    """Controllers UnLocking see :py:meth:`Lo.unlock_controllers() <.utils.lo.Lo.unlock_controllers>`"""
    CONTROLLERS_UNLOCKED = "lo_controllers_unlocked"
    """Controllers UnLocked see :py:meth:`Lo.unlock_controllers() <.utils.lo.Lo.unlock_controllers>`"""

    BRIDGE_DISPOSED = "lo_bridge_disposed"
    """Event when Bridge Component is disposed"""

    RESET = "lo_reset"

    COMPONENT_LOADING = "lo_component_loading"
    """LoInst loading component see :py:meth:`LoInst.load_component() <.utils.inst.lo.lo_inst.LoInst.load_component>`"""
    COMPONENT_LOADED = "lo_component_loaded"
    """LoInst loaded component see :py:meth:`LoInst.load_component() <.utils.inst.lo.lo_inst.LoInst.load_component>`"""

    MACRO_LOADER_ENTER = "lo_macro_loader_enter"
    """MacroLoader enter see :py:class:`~.macro.MacroLoader`"""
    MACRO_LOADER_EXIT = "lo_macro_loader_exit"
    """MacroLoader exit see see :py:class:`~.macro.MacroLoader`"""
