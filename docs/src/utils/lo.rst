.. _utils_lo:

Class Lo
========

.. autoclass:: ooodev.utils.lo.Lo
    :members:
    :undoc-members:

    .. py:property:: Lo.bridge

        Static read-only property

        Gets connection bridge component

        :rtype: XComponent

    .. py:property:: Lo.is_macro_mode

        Static read-only property

        Gets if currently running scripts inside of LO (macro) or standalone

        :rtype: bool

    .. py:property:: Lo.null_date

        Static read-only property

        Gets Value of Null Date in UTC

        :rtype: datetime

        .. note::

            If Lo has no document to determine date from then a
            default date of 1889/12/30 is returned.

    .. py:property:: Lo.star_desktop

        Static read-only property

        Get current desktop

        :rtype: XDesktop

    .. py:property:: Lo.this_component

        Static read-only property

        When the current component is the Basic IDE, the ThisComponent object returns
        in Basic the component owning the currently run user script.
        Above behavior cannot be reproduced in Python.

        When running in a macro this property can be access directly to get the current document.

        When not in a macro then :py:meth:`~.lo.Lo.load_office` must be called first

        :rtype: XComponent

    .. py:property:: Lo.xscript_context

        Static read-only property

        a substitute to `XSCRIPTCONTEXT` LibreOffice/OpenOffice built-in

        :rtype: XScriptContext

    .. py:property:: Lo.bridge_connector

        Static read-only property

        Get the current Bride connection

        :rtype: LoBridgeCommon

    .. py:property:: Lo.current_lo

        Static read-only property

        Get the current Lo instance

        :rtype: LoInst

    .. py:property:: Lo.loader_current

        Static read-only property

        Gets the current ``XComponentLoader`` instance.

        :rtype: XComponentLoader