<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

<xsl:template match="/">
  <office:document-content 
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
    xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
    xmlns:xlink="http://www.w3.org/1999/xlink"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
    xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
    xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
    xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
    xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"
    xmlns:math="http://www.w3.org/1998/Math/MathML"
    xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"
    xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
    xmlns:ooo="http://openoffice.org/2004/office"
    xmlns:ooow="http://openoffice.org/2004/writer"
    xmlns:oooc="http://openoffice.org/2004/calc"
    xmlns:dom="http://www.w3.org/2001/xml-events"
    xmlns:xforms="http://www.w3.org/2002/xforms"
    xmlns:xsd="http://www.w3.org/2001/XMLSchema"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" office:version="1.0">

   <office:automatic-styles>

    <!-- Column styles (co1: column with 6 cm width, 
                        co2: column with 3 cm width) -->
    <style:style style:name="co1" style:family="table-column">
     <style:table-column-properties fo:break-before="auto" style:column-width="6.000cm"/>
    </style:style>
    <style:style style:name="co2" style:family="table-column">
     <style:table-column-properties fo:break-before="auto" style:column-width="3.000cm"/>
    </style:style>

    <!-- Number format styles (N36: date with DD.MM.YYYY,   
                               N107: float with 0,0000) -->
    <number:date-style style:name="N36" number:automatic-order="true">
     <number:day number:style="long"/>
     <number:text>.</number:text>
     <number:month number:style="long"/>
     <number:text>.</number:text>
     <number:year number:style="long"/>
    </number:date-style>

    <number:number-style style:name="N107">
     <number:number number:decimal-places="4" number:min-integer-digits="1"/>
    </number:number-style>


    <!-- Cell styles (ce1: right aligned, 
                      ce2: float with 4 decimal places, 
                      ce3: date) -->
    <style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default">
     <style:table-cell-properties style:text-align-source="fix" style:repeat-content="false"/>
     <style:paragraph-properties fo:text-align="end"/>
    </style:style>

    <style:style style:name="ce2" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N107"/>

    <style:style style:name="ce3" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N36"/>

   </office:automatic-styles>


   <office:body>
    <office:spreadsheet>
     <table:table>

      <!-- Format the first 4 columns of the table -->
      <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
      <table:table-column table:style-name="co2" table:default-cell-style-name="Default"/>
      <table:table-column table:style-name="co2" table:default-cell-style-name="ce2"/>
      <table:table-column table:style-name="co2" table:default-cell-style-name="ce3"/>

      <!-- Insert column labels, first label with default style, the remaining 3 labels right aligned -->
      <table:table-row>
       <table:table-cell><text:p>Purpose</text:p></table:table-cell>
       <table:table-cell table:style-name="ce1" office:value-type="string"><text:p>Amount</text:p></table:table-cell>
       <table:table-cell table:style-name="ce1" office:value-type="string"><text:p>Tax</text:p></table:table-cell>
       <table:table-cell table:style-name="ce1" office:value-type="string"><text:p>Maturity</text:p></table:table-cell>
      </table:table-row>

      <!-- Process XML input: Insert one row for each payment -->
      <xsl:for-each select="payments/payment">
       <table:table-row>

        <!-- Insert string payment purpose -->
        <table:table-cell>
         <text:p><xsl:value-of select="purpose"/></text:p>
        </table:table-cell>

        <!-- Insert float payment amount -->
        <table:table-cell office:value-type="float">
         <xsl:attribute name="office:value"><xsl:value-of select="amount"/></xsl:attribute>
         <text:p><xsl:value-of select="amount"/></text:p>
        </table:table-cell>

        <!-- Insert float payment tax -->
        <table:table-cell office:value-type="float">
         <xsl:attribute name="office:value"><xsl:value-of select="tax"/></xsl:attribute>
         <text:p><xsl:value-of select="tax"/></text:p>
        </table:table-cell>

        <!-- Insert date payment maturity -->
        <table:table-cell office:value-type="date">
         <xsl:attribute name="office:date-value"><xsl:value-of select="maturity"/></xsl:attribute>
         <text:p><xsl:value-of select="maturity"/></text:p>
        </table:table-cell>

       </table:table-row>
      </xsl:for-each>
     </table:table>
    </office:spreadsheet>
   </office:body>
  </office:document-content>
</xsl:template>
</xsl:stylesheet>