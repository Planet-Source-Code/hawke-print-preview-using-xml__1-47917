<html xmlns:xsl="http://www.w3.org/TR/WD-xsl">
<body>
<table STYLE="font-family:Arial; font-size:smaller" bordercolor="black" cellspacing="0" cellpadding="3" border="0">
  <tr bgcolor="black">
    <th><font color="white">Account Name</font></th>
    <th><font color="white">Balance</font></th>
    <th><font color="white">Comments</font></th>
  </tr>
<xsl:for-each select="xml/rs:data/z:row">
  <tr>
    <td><xsl:value-of select="@AccountName"/></td>
    <td><xsl:value-of select="@Balance"/></td> 
    <td><xsl:value-of select="@Comments"/></td>
  </tr>
</xsl:for-each>
</table>
</body>
</html>

