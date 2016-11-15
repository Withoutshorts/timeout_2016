<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/xsl/transform">
<xsl:variable name="config" select="'conf.xml'"/>

<xsl:output method="xml" version="1.0" encoding="ISO-8859-1" omit-xml-declaration="no" indent="no" media-type="text/html"/>

<xsl:template match="conf">
<html><head><title>output</title></head>
<body>
<xsl:for-each select="modul">
	<p>Her:<xsl:value-of select="."/></p>
</xsl:for-each>
</body></html>
</xsl:template>
</xsl:stylesheet>