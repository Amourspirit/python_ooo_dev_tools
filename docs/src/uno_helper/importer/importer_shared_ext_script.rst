.. _ns_uno_helper_importer_shared_ext_script:

Module importer_shared_ext_script
=================================

Function: importer_shared_ext_script
------------------------------------

.. function:: importer_shared_ext_script(ext_name)

    A context manager that manages adding ``ImporterSharedExtScript`` to the ``sys.meta_path``.

    This context manager ensures that the resource is properly acquired and released.

    The importer will only search in the ``python/scripts`` folder of the extension.

    Args:
        ext_name (str): The name of the extension that contains the  module to be imported.

    Returns:
        None:

    Note:
        The extension module must be in registered in LibreOffice as a shared extension.

    Example:

        In this example the context manager is used to import a module from a Calc document.

        .. code-block:: python

            from __future__ import annotations
            from pathlib import Path
            from ooodev.calc import CalcDoc
            from ooodev.loader import Lo
            from ooodev.uno_helper.importer import importer_shared_ext_script


            def main():
                _ = Lo.load_office(connector=Lo.ConnectPipe())
                doc = None
                try:
                    doc = CalcDoc.create_doc()

                    with importer_user_ext_script("apso"):
                        import tools

                    print("Done")

                except Exception as e:
                    print(f"Error: {e}")

                finally:
                    if doc:
                        doc.close()
                    Lo.close_office()


            if __name__ == "__main__":
                main()

        In this example, a ``ImporterSharedExtScript`` instance is automatically managed, ensuring proper cleanup.


class ImporterSharedExtScript
-----------------------------

.. autoclass:: ooodev.uno_helper.importer.ImporterSharedExtScript
    :members:
    :undoc-members:
    :show-inheritance:


