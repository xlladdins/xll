<?xml version="1.0" encoding="iso-8859-1"?>
<xsl:transform xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
               version="2.0"
>
  <xsl:output method="html" encoding="iso-8859-1" version="4.0" doctype-public="-//W3C//DTD HTML 4.01//EN" indent="no"/>
  <xsl:template match="/data/*">
    <html lang="en">
      <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
        <link href="xll.css" rel="stylesheet" type="text/css" />
        <title>
          <xsl:value-of select="title"/>
        </title>
      </head>
      <body>
        <h1>
          <xsl:value-of select="functionText"/>
        </h1>
      </body>
    </html>
  </xsl:template>
</xsl:transform>