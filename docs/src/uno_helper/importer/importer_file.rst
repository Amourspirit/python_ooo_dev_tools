.. _ns_uno_helper_importer_file:

Module importer_file
====================

Function: importer_file
-----------------------

.. function:: importer_file(module_path)

    A context manager that manages adding ``ImporterFile`` to the ``sys.meta_path``.

    This context manager ensures that the resource is properly acquired and released.

    Args:
        module_path: Path to the folder where the module exist for import.

    Returns:
        None:

    Example:

        In this example the context manager is used to import a module from a known path.

        .. code-block:: python

            from __future__ import annotations
            from pathlib import Path
            from ooodev.calc import CalcDoc
            from ooodev.loader import Lo
            from ooodev.uno_helper.importer import importer_file


            def main():
                _ = Lo.load_office(connector=Lo.ConnectPipe())
                doc = None
                try:
                    fnm = Path.cwd() / "mods"
                    doc = CalcDoc.create_doc()

                    with importer_file(fnm):
                        import big_worker

                    big_worker.work()

                    print("Done")

                except Exception as e:
                    print(f"Error: {e}")

                finally:
                    if doc:
                        doc.close()
                    Lo.close_office()


            if __name__ == "__main__":
                main()

        In this example, a ``ImporterFile`` instance is automatically managed, ensuring proper cleanup.


class ImporterFile
------------------

.. autoclass:: ooodev.uno_helper.importer.ImporterFile
    :members:
    :undoc-members:
    :show-inheritance:


