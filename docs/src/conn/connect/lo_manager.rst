.. _conn_connect_lo_manager:

Class LoManager
===============

LoManager is low level content manager that connects to LibreOffice and then disconnects
when the block has been executed.

Generally speaking this Manager is not needed when using |odev|.

Example

    .. code-block:: python

        from ooodev.conn.connect import LoManager

        with LoManager() as mgr:
            smgr = mgr.ctx.getServiceManager()
            desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", mgr.ctx)
            # other processing

.. seealso::

    - :py:class:`.Lo.Loader`
    - :ref:`ch02`


.. autoclass:: ooodev.conn.connect.LoManager
    :members:
