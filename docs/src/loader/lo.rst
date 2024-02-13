.. _ooodev.loader.Lo:

Class Lo
========

.. autoclass:: ooodev.loader.Lo
    :members:
    :undoc-members:

    .. py:property:: Lo.bridge

        Static read-only property.

        Gets connection bridge component

        |lo_unsafe|

        :rtype: XComponent

    .. py:property:: Lo.current_doc

        Static read-only property.

        Gets the current document. Such as ``ooodev.calc.CalcDoc`` or ``ooodev.write.WriteDoc``.

        |lo_unsafe|

        :rtype: OfficeDocumentT

    .. py:property:: Lo.is_macro_mode

        Static read-only property.

        Gets if currently running scripts inside of LO (macro) or standalone.

        :rtype: bool

    .. py:property:: Lo.null_date

        Static read-only property.

        Gets Value of Null Date in UTC.

        |lo_safe|

        :rtype: datetime

        .. note::

            If Lo has no document to determine date from then a
            default date of 1889/12/30 is returned.

    .. py:property:: Lo.star_desktop

        Static read-only property.

        Get current desktop.

        |lo_unsafe|

        :rtype: XDesktop

    .. py:property:: Lo.this_component

        Static read-only property.

        When the current component is the Basic IDE, the ThisComponent object returns
        in Basic the component owning the currently run user script.
        Above behavior cannot be reproduced in Python.

        When running in a macro this property can be access directly to get the current document.

        When not in a macro then :py:meth:`~.lo.Lo.load_office` must be called first.

        |lo_unsafe|

        :rtype: XComponent

    .. py:property:: Lo.xscript_context

        Static read-only property a substitute to `XSCRIPTCONTEXT` LibreOffice/OpenOffice built-in.

        |lo_unsafe|

        :rtype: XScriptContext

    .. py:property:: Lo.bridge_connector

        Static read-only property.

        |lo_unsafe|

        Get the current Bride connection

        :rtype: LoBridgeCommon

    .. py:property:: Lo.current_lo

        Static read-only property.

        |lo_unsafe|

        Get the current Lo instance

        :rtype: LoInst

    .. py:property:: Lo.loader_current

        Static read-only property.

        |lo_unsafe|

        Gets the current ``XComponentLoader`` instance.

        :rtype: XComponentLoader
    
    .. py:property:: desktop

        Static read-only property.

        Get the current Desktop instance.

        |lo_unsafe|

        Desktop instance. Component property implments is ``XDesktop``

        :rtype: TheDesktop
    

    .. py:property:: global_event_broadcaster

        Static read-only property.

        Get the current Global Broadcaster instance.
        This is a simpler way to add listeners to the global event broadcaster that are broadcasted by ``theGlobalEventBroadcaster``.

        |lo_unsafe|

        Desktop instance. Component property implments is ``XDesktop``

        :rtype: TheGlobalEventBroadcaster

        .. code-block:: python

            import contextlib
            from typing import TYPE_CHECKING, Any, cast
            from ooodev.loader import Lo

            if TYPE_CHECKING:
                from com.sun.star.document import DocumentEvent
            else:
                DocumentEvent = Any

            #  Add a listener to the global event broadcaster
            Lo.global_event_broadcaster.add_event_document_event_occurred(_on_global_document_event)

            def _on_global_document_event(src: Any, event: EventArgs, *args, **kwargs) -> None:
                # This is a listener for the global event broadcaster
                with contextlib.suppress(Exception):
                    doc_event = cast(DocumentEvent, event.event_data)
                    name = doc_event.EventName
                    if name == "OnUnfocus":
                        # only interested in the OnUnfocus event
                        self._clear_cache()
        
        .. seealso::

            - :ref:`the_global_event_broadcaster`
            - `LibreOffice API - theGlobalEventBroadcaster <https://api.libreoffice.org/docs/idl/ref/singletoncom_1_1sun_1_1star_1_1frame_1_1theGlobalEventBroadcaster.html>`__