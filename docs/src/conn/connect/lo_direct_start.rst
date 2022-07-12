Class LoDirectStart
===================

This class is used for connecting to LibreOffice via macros.

Macros do not require a connection as they already have access to LibreOffice API by their nature.

Generally speaking this class will not be called directly. It is used internally by :py:class:`~.utils.lo.Lo`.

Lo macro Example:

    .. include:: ../../../resources/utils/lo_connect_direct_ex.rst

.. note::

    When script is running in macro is is not necessary to call :py:meth:`.Lo.load_office`.
    In a macro use ``Lo.XSCRIPTCONTEXT`` and/or ``Lo.ThisComponent``

.. seealso:: 

    - :ref:`ch02`

.. autoclass:: ooodev.conn.connect.LoDirectStart
    :members:
    :inherited-members:
    :show-inheritance:
