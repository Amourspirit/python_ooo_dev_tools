<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
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
  xmlns:math="http://www.w3.org/1998/Math/MathML"
  xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"
  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
  office:version="1.0">
<xsl:template match="/">
	<office:document-content>
		<office:body>
			<office:text>
				<xsl:apply-templates select="club-database/association"/>
			</office:text>
		</office:body>
	</office:document-content>
</xsl:template>

<xsl:template match="association">
	<text:h text:outline-level="1" text:style-name="Association">
		<xsl:value-of select="@id"/>
	</text:h>
	<xsl:apply-templates select="club"/>
</xsl:template>

<xsl:template match="club">
	<text:h text:outline-level="2" text:style-name="Club Name">
		<xsl:value-of select="name" />
		<xsl:text> </xsl:text>
		<text:span text:style-name="Club Code"><xsl:value-of
			select="@id" /></text:span>
	</text:h>
	<text:p text:style-name="Default">
		<xsl:text>Chartered: </xsl:text>
		<text:span text:style-name="Charter">
			<xsl:value-of select="@charter"/>
		</text:span>
	</text:p>
	<text:p text:style-name="Default">
		<xsl:text>Contact: </xsl:text>
		<text:span text:style-name="Contact">
			<xsl:value-of select="contact"/>
		</text:span>
	</text:p>
	<text:p text:style-name="Default">
		<xsl:text>Location: </xsl:text>
		<text:span text:style-name="Location">
			<xsl:value-of select="location"/>
		</text:span>
	</text:p>
	<text:p text:style-name="Default">
		<xsl:text>Phone: </xsl:text>
		<text:span text:style-name="Phone">
			<xsl:value-of select="phone"/>
		</text:span>
	</text:p>

	<xsl:choose>
		<xsl:when test="count(email) = 1">
			<text:p text:style-name="Default">
				<xsl:text>Email: </xsl:text>
				<text:span text:style-name="Email">
					<xsl:value-of select="email"/>
				</text:span>
			</text:p>
		</xsl:when>
		<xsl:when test="count(email) &gt; 1">
			<text:p text:style-name="Default">
				<text:span>Email:</text:span>
			</text:p>
			<text:list text:style-name="UnorderedList">
				<xsl:for-each select="email">
					<text:list-item>
						<text:p text:style-name="Default">
							<text:span text:style-name="Email">
								<xsl:value-of select="."/>
							</text:span>
						</text:p>
					</text:list-item>
				</xsl:for-each>
			</text:list>
		</xsl:when>
	</xsl:choose>

	<xsl:apply-templates select="age-groups"/>
	
	<xsl:apply-templates select="info"/>
</xsl:template>

<xsl:template match="age-groups">
	<text:p text:style-name="Default">
		<xsl:text>Age Groups: </xsl:text>
		<text:span text:style-name="Age Groups">
			<xsl:if test="contains(@type,'K')">
				<xsl:text>Kids </xsl:text>
			</xsl:if>
			<xsl:if test="contains(@type,'C')">
				<xsl:text>Cadets </xsl:text>
			</xsl:if>
			<xsl:if test="contains(@type,'J')">
				<xsl:text>Juniors </xsl:text>
			</xsl:if>
			<xsl:if test="contains(@type,'O')">
				<xsl:text>Open </xsl:text>
			</xsl:if>
			<xsl:if test="contains(@type,'W')">
				<xsl:text>Women </xsl:text>
			</xsl:if>
		</text:span>
	</text:p>
</xsl:template>

<xsl:template match="info">
	<text:p text:style-name="Club Info">
		<xsl:if test="normalize-space(.) != ''">
			<xsl:apply-templates/>
		</xsl:if>
	</text:p>
</xsl:template>

<xsl:template match="a">
	<text:a xlink:type="simple" xlink:href="{@href}"><xsl:value-of select="."/></text:a>
</xsl:template>

</xsl:stylesheet>
