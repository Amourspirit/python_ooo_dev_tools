.. code-block:: python

    from ooodev.loader.lo import Lo
    from ooodev.office.calc import Calc

    loader = Lo.load_office(connector=Lo.ConnectPipe())
    doc = Calc.create_doc(loader)
    # do work ...
    Lo.close_doc(doc=doc)
    Lo.close_office()
