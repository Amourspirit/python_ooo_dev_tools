.. tabs::

    .. tab:: Python

        .. tabs::

            .. tab:: Lo Class
            
                .. code-block:: python
                
                    def main() -> None:
                        loader = Lo.load_office(Lo.ConnectSocket(headless=True))
                        doc = Lo.create_doc(doc_type=Lo.DocTypeStr.WRITER, loader=loader)

                        # use the Office API to manipulate doc...

                        Lo.save_doc(doc, "foo.docx") # save as a Word file
                        Lo.close_doc(doc)
                        lo.close_office()
            
            .. tab:: Write Class

                .. code-block:: Python
                
                    def main() -> None:
                        loader = Lo.load_office(Lo.ConnectSocket(headless=True))
                        doc = Write.create_doc(loader=loader)

                        # use the Office API to manipulate doc...

                        Write.save_doc(doc, "foo.docx") # save as a Word file
                        Lo.close_doc(doc)
                        lo.close_office()
            
            .. tab:: Calc Class

                .. code-block:: Python
                
                    def main() -> None:
                        loader = Lo.load_office(Lo.ConnectSocket(headless=True))
                        doc = Calc.create_doc(loader=loader)
                        sheet = Calc.get_sheet(doc=doc, index=0)

                        # use the Office API to manipulate doc...

                        Calc.save_doc(doc, "foo.ods")
                        Lo.close_doc(doc)
                        lo.close_office()