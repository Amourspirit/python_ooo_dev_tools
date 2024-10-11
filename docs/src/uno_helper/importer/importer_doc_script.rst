.. _ns_uno_helper_importer_doc_script:

Module ImporterDocScript
========================

Function: importer_doc_script
-----------------------------

.. function:: importer_doc_script()

    A context manager that manages adding ``ImporterUserScript`` to the ``sys.meta_path``.

    This context manager ensures that the resource is properly acquired and released.

    Returns:
        None:


    Example:

        In this example the context manager is used to import a module from a Calc document.

        .. code-block:: python

            from __future__ import annotations
            from pathlib import Path
            from ooodev.calc import CalcDoc
            from ooodev.loader import Lo
            from ooodev.uno_helper.importer import importer_doc_script


            def main():
                _ = Lo.load_office(connector=Lo.ConnectPipe())
                doc = None
                try:
                    fnm = Path.cwd() / "calc_runner.ods"
                    doc = CalcDoc.open_doc(fnm=fnm, visible=True)

                    with importer_doc_script():
                        import mod_hello

                    mod_hello.say_hello()

                    print("Done")

                except Exception as e:
                    print(f"Error: {e}")

                finally:
                    if doc:
                        doc.close()
                    Lo.close_office()


            if __name__ == "__main__":
                main()

        In this example, a ``ImporterUserScript`` instance is automatically managed, ensuring proper cleanup.


class ImporterDocScript
------------------------

.. autoclass:: ooodev.uno_helper.importer.ImporterDocScript
    :members:
    :undoc-members:
    :show-inheritance:


