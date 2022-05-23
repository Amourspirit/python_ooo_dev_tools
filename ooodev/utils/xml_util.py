# coding: utf-8
# Python conversion of XML.java by Andrew Davison, ad@fivedots.coe.psu.ac.th
# See Also: https://fivedots.coe.psu.ac.th/~ad/jlop/
from typing import Iterable, Union, Tuple, List, overload
from xml.dom import minidom
import urllib.request
from xml.dom.minicompat import NodeList
from lxml import etree as ET
from . import lo as mLo
from .gen_util import TableHelper

_xml_parser = ET.XMLParser(remove_blank_text=True)


class XML:
    @classmethod
    def load_doc(cls, fnm: str) -> Union[minidom.Document, None]:
        """
        Gets a document from a file

        Args:
            fnm (str): XML file to load.

        Returns:
            Document | None: XML Document on successful load; Otherwise, None.
        """
        try:
            with open(fnm) as file:
                elem = ET.XML(text=file.read(), parser=_xml_parser)
            doc = minidom.parseString(string=ET.tostring(elem))
            return doc
        except Exception as e:
            print(e)
        return None

    @classmethod
    def url_2_doc(cls, url: str) -> Union[minidom.Document, None]:
        """
        Gets a XML Document from remote souce.

        Args:
            url (str): Url for a remote XML Document

        Returns:
            Document | None: XML Document on successful load; Otherwise, None.
        """
        try:
            with urllib.request.urlopen(url) as url_data:

                elem = ET.XML(text=url_data.read().decode(), parser=_xml_parser)
                doc = minidom.parseString(string=ET.tostring(elem))
                return doc
        except Exception as e:
            print(e)
        return None

    @classmethod
    def str_2_doc(cls, xml_str: str) -> Union[minidom.Document, None]:
        """
        Gets a XML document from xml string.

        Args:
            xml_str (str): XML string.

        Returns:
            Document | None: XML Document on successful load; Otherwise, None.
        """
        try:
            elem = ET.XML(text=xml_str, parser=_xml_parser)
            doc = minidom.parseString(string=ET.tostring(elem))
            return doc
        except Exception as e:
            print(e)
        return None

    @staticmethod
    def save_doc(doc: minidom.Document, xml_fnm: str) -> None:
        """
        Save doc to xml file.

        Args:
            doc (Document): doc to save.
            xml_fnm (str): Output file path.
        """
        try:
            with open(xml_fnm, "w") as file:

                sx: str = doc.toprettyxml(indent="  ")
                # remove any empty lines, there if often a lot with toprettyxml()
                lines = [line for line in sx.splitlines() if line.strip() != ""]
                lines.append("")  # end with empty line
                clean_sx = "\n".join(lines)
                file.write(clean_sx)
                # doc.writexml(writer=file, indent="  ")
        except Exception as e:
            print(f"Unable to save document to {xml_fnm}")
            print(f"  {e}")

    # --------------- DOM data extraction -----------------------

    @staticmethod
    def get_node(tag_name: str, nodes: NodeList) -> Union[minidom.Node, None]:
        """
        Gets the fist tag_name found in nodes.

        Args:
            tag_name (str): tag name to find in nodes.
            nodes (NodeList): Nodes to search

        Returns:
            Node | None: First found node; Othwewise, None
        """
        name = tag_name.lower()
        for node in nodes:
            if node.tagName.lower() == name:
                return node
        return None

    @overload
    @staticmethod
    def get_node_value(node: minidom.Node) -> str:
        """
        Get the text stored in the node

        Args:
            node (Node): Node to get value of.

        Returns:
            str: Node value.
        """
        ...

    @overload
    @staticmethod
    def get_node_value(tag_name: str, nodes: NodeList) -> str:
        """
        Gets firt tag_name node in the list and returns it text.

        Args:
            tag_name (str): tag_name to search for.
            nodes (NodeList): List of nodes to search.

        Returns:
            str: Node value if found; Otherwise empty str.
        """
        ...

    @classmethod
    def get_node_value(cls, *args, **kwargs) -> str:
        ordered_keys = ("first", "second")
        kargs = {}
        if "node" in kwargs:
            kargs["first"] = kwargs["node"]
        elif "tag_name" in kwargs:
            kargs["first"] = kwargs["tag_name"]
        elif "nodes" in kwargs:
            kargs["second"] = kwargs["nodes"]
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg
        k_len = len(kargs)

        if k_len == 1:
            return cls._get_node_val(kargs["first"])
        if k_len == 2:
            return cls._get_node_val2(kargs["first"], kargs["second"])
        raise ValueError("Incorrect number of prameter supplied:")

    @staticmethod
    def _get_node_val(node: minidom.Node) -> str:
        if node is None:
            return ""
        if not node.hasChildNodes():
            return ""
        child_nodes: NodeList = node.childNodes
        for node in child_nodes:
            if node.nodeType == minidom.Node.TEXT_NODE:
                return str(node.data).strip()
        return ""

    @classmethod
    def _get_node_val2(cls, tag_name: str, nodes: NodeList) -> str:
        if nodes is None:
            return ""
        name = tag_name.lower()
        for node in nodes:
            if node.nodeName.lower() == name:
                return cls.get_node_value(node)
        return ""

    @classmethod
    def get_node_values(cls, nodes: NodeList) -> Tuple[str, ...]:
        """
        Gets all the node values

        Args:
            nodes (NodeList): Nodes to get values of.

        Returns:
            Tuple[str, ...]: Node Values
        """
        vals = []
        for node in nodes:
            val = cls.get_node_value(node)
            if val != "":
                vals.append(val)
        return tuple(val)

    @staticmethod
    def get_node_attr(attr_name: str, node: minidom.Node) -> str:
        """
        Get the named attribute value from node

        Args:
            attr_name (str): Attribute Name
            node (minidom.Node): Node to get attribue of.

        Returns:
            str: Attribute value if found; Othwewise empty str.
        """
        if node is None:
            return ""
        # attrs is {} if there are no attributes
        attrs = dict(node.attributes.items())
        name = attr_name.lower()
        for k, v in attrs.items():
            if str(k).lower() == name:
                return str(v)
        return ""

    @classmethod
    def get_all_node_values(cls, row_nodes: NodeList, col_ids: Iterable[str]) -> List[list]:
        """
        assumes an XML structure like
        ::
            <root>
                <rowID>
                    <col1ID>str1</col1ID>
                    <col2ID>str2</col2ID>
                    <col3ID>str2</col3ID>
                    <col4ID>str2</col4ID>
                </rowID>
                <rowID>
                    <col1ID>row2-1</col1ID>
                    <col2ID>row2-2</col2ID>
                    <col3ID>row2-3</col3ID>
                    <col4ID>row2-4</col4ID>
                </rowID>
                <rowID>
                    <col1ID>row3-1</col1ID>
                    <col2ID>row3-2</col2ID>
                    <col3ID>row3-3</col3ID>
                    <col4ID>row3-4</col4ID>
                </rowID>
            </root>

        The data from a sequence of <col>s becomes one row in the
        generated 2D array.

        The first row of the 2D array contains the col ID strings.

        Note:
            col_ids must match the column names:

            colids = ["col1ID", "col2ID", "col3ID", "col4ID"]
        """
        num_rows = len(row_nodes)
        num_cols = len(col_ids)
        if num_cols == 0 or num_rows == 0:
            return []
        data = TableHelper.make_2d_array(num_rows=num_rows, num_cols=num_cols)
        # data = [[1] * num_cols for _ in range(num_rows + 1)]
        # put column strings in first row of list
        for col, _ in enumerate(col_ids):
            data[0][col] = mLo.Lo.capitalize(col_ids[col])

        for i, node in enumerate(row_nodes):
            # extract all the column strings for ith row
            col_nodes = node.childNodes
            for col in range(num_cols):
                data[i + 1][col] = cls.get_node_value(col_ids[col], col_nodes)
        return data

    # ------------------------- XLS transforming ----------------------

    @staticmethod
    def apply_xslt(xml_fnm: str, xls_fnm: str) -> Union[str, None]:
        """
        Transforms xml file using XLST

        Args:
            xml_fnm (str): XML source file path.
            xls_fnm (str): XSL source file path.

        Returns:
            Union[str, None]: String of XML that has been transformed.
        """
        try:
            print(f"Applying filter '{xls_fnm}' to '{xml_fnm}'")
            dom = ET.parse(xml_fnm, parser=_xml_parser)
            xslt = ET.parse(xls_fnm)
            transform = ET.XSLT(xslt)
            newdom = transform(dom)
            t_result = ET.tostring(newdom, encoding="unicode")  # unicode produces string
            return t_result
        except Exception as e:
            print(f"Unable to transform '{xml_fnm}' with '{xls_fnm}'")
            print(f"    {e}")
        return None

    @staticmethod
    def apply_xslt_2_str(xml_str: str, xls_fnm: str) -> Union[str, None]:
        """
        Transforms xml using XLST

        Args:
            xml_str (str): XML string.
            xls_fnm (str): XSL source file path.

        Returns:
            Union[str, None]: String of XML that has been transformed.
        """
        try:
            print(f"Applying the filter in '{xls_fnm}'")
            dom = ET.fromstring(xml_str)
            xslt = ET.parse(xls_fnm, parser=_xml_parser)

            transform = ET.XSLT(xslt)
            newdom = transform(dom)
            t_result = ET.tostring(newdom, encoding="unicode")  # unicode produces string
            return t_result
        except Exception as e:
            print("Unable to transform the string")
            print(f"    {e}")
        return None

    @staticmethod
    def indent(xml_fnm: str) -> Union[str, None]:
        """
        Indents xml

        Args:
            xml_fnm (str): xml file path.

        Returns:
            Union[str, None]: Indented xml on success; Otherwise, None
        """
        try:
            dom = ET.parse(xml_fnm, parser=_xml_parser)
            result = ET.tostring(dom, encoding="unicode", pretty_print=True)
            return result
        except Exception as e:
            print(f"Unable to indent '{xml_fnm}'")
            print(f"    {e}")
        return None

    @staticmethod
    def indent_2_str(xml_str: str) -> Union[str, None]:
        """
        Indents xml

        Args:
            xml_str (str): xml string to indent

        Returns:
            Union[str, None]: Indented xml on success; Otherwise, None
        """
        try:
            dom = ET.fromstring(xml_str, parser=_xml_parser)
            result = ET.tostring(dom, encoding="unicode", pretty_print=True)
            return result
        except Exception as e:
            print("Unable to indent the xml string")
            print(f"    {e}")
        return None

    @staticmethod
    def get_flat_fiter_name(doc_type: str) -> str:
        """
        Gts the Flat XML filter name for the doc type.

        Args:
            doc_type (str): Document type.

        Returns:
            str: Flat XML filter name.
        """
        if doc_type == mLo.Lo.WRITER_STR:
            return "OpenDocument Text Flat XML"
        elif doc_type == mLo.Lo.CALC_STR:
            return "OpenDocument Spreadsheet Flat XML"
        elif doc_type == mLo.Lo.DRAW_STR:
            return "OpenDocument Drawing Flat XML"
        elif doc_type == mLo.Lo.IMPRESS_STR:
            return "OpenDocument Presentation Flat XML"
        else:
            print("No Flat XML filter for this document type; using Flat text")
            return "penDocument Text Flat XML"
