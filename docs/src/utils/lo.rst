Class Lo
========

.. autoclass:: ooodev.utils.lo.Lo
    :members:
    :undoc-members:

    .. py:property:: Lo.bridge

        Static readonly property

        Gets connection bridge component

        :rtype: XComponent

    .. py:property:: Lo.is_macro_mode

        Static readonly property

        Gets if currently running scripts inside of LO (macro) or standalone

        :rtype: bool

    .. py:property:: Lo.null_date

        Static readonly property

        Gets Value of Null Date in UTC

        :rtype: datetime

        .. note::

            If Lo has no document to determine date from then a
            default date of 1889/12/30 is returned.

    .. py:property:: Lo.star_desktop

        Static readonly property

        Get current desktop

        :rtype: XDesktop

    .. py:property:: Lo.this_component

        Static readonly property

        When the current component is the Basic IDE, the ThisComponent object returns
        in Basic the component owning the currently run user script.
        Above behaviour cannot be reproduced in Python.

        When running in a macro this property can be access directly to get the current document.

        When not in a macro then :py:meth:`~.lo.Lo.load_office` must be called first

        :rtype: XComponent

    .. py:property:: Lo.xscript_context

        Static readonly property

        a substitute to `XSCRIPTCONTEXT` (Libre|Open)Office built-in

        :rtype: XScriptContext