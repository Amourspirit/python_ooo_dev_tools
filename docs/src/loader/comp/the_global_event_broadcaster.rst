.. _the_global_event_broadcaster:

Class TheGlobalEventBroadcaster
===============================

Introduction
------------

This is a simpler way to add listeners to the global event broadcaster that are broadcasted by ``theGlobalEventBroadcaster``.

.. seealso::

    - :ref:`ooodev.loader.Lo`
    - `LibreOffice API - theGlobalEventBroadcaster <https://api.libreoffice.org/docs/idl/ref/singletoncom_1_1sun_1_1star_1_1frame_1_1theGlobalEventBroadcaster.html>`__

Example Usage
-------------

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

Class Declaration
-----------------

.. autoclass:: ooodev.loader.comp.the_global_event_broadcaster.TheGlobalEventBroadcaster
    :members:
    :undoc-members:
    :show-inheritance:
    :inherited-members:
