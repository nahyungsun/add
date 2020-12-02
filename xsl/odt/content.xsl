<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
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
  xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"
  xmlns:ooo="http://openoffice.org/2004/office"
  xmlns:ooow="http://openoffice.org/2004/writer"
  xmlns:oooc="http://openoffice.org/2004/calc"
  xmlns:dom="http://www.w3.org/2001/xml-events"
  xmlns:xforms="http://www.w3.org/2002/xforms"
  xmlns:xsd="http://www.w3.org/2001/XMLSchema"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:rpt="http://openoffice.org/2005/report"
  xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"
  xmlns:xhtml="http://www.w3.org/1999/xhtml"
  xmlns:grddl="http://www.w3.org/2003/g/data-view#"
  xmlns:tableooo="http://openoffice.org/2009/table"
  xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0"
  xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0"
  xmlns:css3t="http://www.w3.org/TR/css3-text/"
  >
  <xsl:import href="common.xsl" />
  <xsl:output method="xml" encoding="utf-8" indent="no" />

  <xsl:template match="/">
    <xsl:apply-templates mode="office:document-content" select="HwpDoc" />
  </xsl:template>

  <xsl:template mode="office:document-content" match="HwpDoc">
    <office:document-content
      office:version="1.2"
      grddl:transformation="http://docs.oasis-open.org/office/1.2/xslt/odf2rdf.xsl">
      <office:scripts/>
      <office:font-face-decls/>
      <xsl:apply-templates mode="office:automatic-styles" select="BodyText" />
      <xsl:apply-templates mode="office:body" select="BodyText" />
    </office:document-content>
  </xsl:template>

  <xsl:template mode="office:automatic-styles" match="BodyText">
    <office:automatic-styles>
      <xsl:apply-templates mode="style-style-for-paragraph-and-text" select="//Paragraph" />
      <xsl:apply-templates mode="style:style" select="SectionDef//TableControl" />
      <xsl:apply-templates mode="style-style-for-table-cells" select="SectionDef//TableControl" />
      <xsl:apply-templates mode="style:style" select="SectionDef//ShapeComponent" />
    </office:automatic-styles>
  </xsl:template>

</xsl:stylesheet>
