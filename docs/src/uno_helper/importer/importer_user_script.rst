.. _ns_uno_helper_importer_user_script:

Module importer_user_script
===========================

Function: importer_user_script
--------------------------------

.. function:: importer_user_script(ext_name)

    A context manager that manages adding ``ImporterUserScript`` to the ``sys.meta_path``.

    This context manager ensures that the resource is properly acquired and released.

    The importer will only search in the ``Scripts/python`` folder of the LibreOffice User Libraries.
    On Linux this is typically ``~/.config/libreoffice/4/user/Scripts/python``.

    Returns:
        None:

    Example:

        In this example the context manager is used to import a module from the user library.

        .. code-block:: python

            from __future__ import annotations
            from pathlib import Path
            from ooodev.calc import CalcDoc
            from ooodev.loader import Lo
            from ooodev.uno_helper.importer import importer_user_script


            def main():
                _ = Lo.load_office(connector=Lo.ConnectPipe())
                doc = None
                try:
                    doc = CalcDoc.create_doc()

                    with importer_user_script():
                        import my_module

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


class ImporterUserScript
--------------------------

.. autoclass:: ooodev.uno_helper.importer.ImporterUserScript
    :members:
    :undoc-members:
    :show-inheritance:


