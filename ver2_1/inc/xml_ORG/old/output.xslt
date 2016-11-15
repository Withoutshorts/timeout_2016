<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/xsl/transform">
<xsl:output method="xml" indent="yes"/>
<xsl:template match="modul">
<html><head><title>output</title></head>
<body>
<!--<xsl:for-each select="/conf/modul::*">-->
	<p><xsl:value-of select="."/></p>
<!--</xsl:for-each>-->
</body></html>
</xsl:template>
</xsl:stylesheet>