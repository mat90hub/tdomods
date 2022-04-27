#!/usr/bin/env tclsh

# https://stackoverflow.com/questions/7908627/how-to-apply-an-xslt-transformation-that-includes-spaces-to-an-xml-doc-using-tdo

set DRV [string trim {
<?xml version="1.0" encoding="UTF-8"?>
<definitions devices="myDevice">
    <reg offset="0x0000" mnem="someRegister">
        <field mnem="someField" msb="31" lsb="24" />        
    </reg>
</definitions>
}]

set XSLT [string trim {
<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:template match="/">
<xsl:for-each select="definitions/reg">
<xsl:text>#define </xsl:text>
<xsl:value-of select="translate(@mnem,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')"/>
<xsl:text xml:space="preserve"> </xsl:text>
<xsl:value-of select="@offset"/>
</xsl:for-each>
</xsl:template>
</xsl:stylesheet>
}]


package require tdom
set IN [dom parse $DRV]
set xslt [dom parse -keepEmpties $XSLT]
set OUT [$IN xslt $xslt]
$OUT asText


