.. code-block:: python

    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    loader = Lo.load_office(connector=Lo.ConnectSocket())
    doc = Calc.create_doc(loader)
    # do work ...
    Lo.close_doc(doc=doc)
    Lo.close_office()
