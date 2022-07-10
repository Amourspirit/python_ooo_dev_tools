Class LoManager
===============

LoManager is low level contenxt manager that connects to LibreOffice and then disconnects
when the block has been executed.

Generally speaking this Manager is not needed when using |app_name_short|.

Example

    .. code-block:: python

        from ooodev.conn.connect import LoManager

        with LoManager() as mgr:
            smgr = mgr.ctx.getServiceManager()
            desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", mgr.ctx)
            # other processing

.. seealso::

    :py:class:`.Lo.Loader`

.. autoclass:: ooodev.conn.connect.LoManager
    :members:
