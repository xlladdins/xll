<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:output method="html"
              encoding="utf-8"
              version="4.0"
              doctype-public="-//W3C//DTD HTML 4.01//EN"
              omit-xml-declaration="yes"
              indent="no"/>
  <xsl:template match="/data">
    <html lang="en">
      <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
        <link href="xll.css" rel="stylesheet" type="text/css" />
        <link rel="stylesheet"
              href="https://cdn.jsdelivr.net/npm/katex@0.12.0/dist/katex.min.css"
              integrity="sha384-AfEj0r4/OFrOo5t7NnNe46zW/tFgW6x/bCJG8FqQCEo3+Aro6EYUG4+cU+KJWu/X"
              crossorigin="anonymous" />
        <script defer=""
                src="https://cdn.jsdelivr.net/npm/katex@0.12.0/dist/katex.min.js"
                integrity="sha384-g7c+Jr9ZivxKLnZTDUhnkOnsh30B4H0rpLUpJ4jAIKs4fnJI+sEnkvrMWph2EDg4"
                crossorigin="anonymous">
        </script>
        <script defer=""
                src="https://cdn.jsdelivr.net/npm/katex@0.12.0/dist/contrib/auto-render.min.js"
                integrity="sha384-mll67QQFJfxn0IYznZYonOWZ644AWYC+Pt2cHqMaRhXVrursRwvLnLaebdGIlYNa"
                crossorigin="anonymous"
                onload="renderMathInElement(document.body, {fleqn: true});">
        </script>
        <title>
          <xsl:value-of select="title"/>
        </title>
      </head>
      <body>
        <h1>
          <xsl:value-of select="functionText"/>
          <xsl:value-of select="type"/>
        </h1>
        <p>
          This article describes the formula syntax of the
          <xsl:value-of select="functionText"/>
          <xsl:value-of select="type"/>
        </p>
        <h2>Description</h2>
        <p>
          <xsl:value-of select="functionHelp"/>
        </p>
        <h2>Syntax</h2>
        <p>
          <xsl:value-of select="functionText"/>(<xsl:value-of select="argumentText"/>)
        </p>
        <blockquote>
          <table>
            <tbody>
              <xsl:for-each select="arguments/arg">
                <tr>
                  <td>
                    <xsl:value-of select="name"/>
                  </td>
                  <td>
                    <xsl:value-of select="help"/>
                  </td>
                </tr>
              </xsl:for-each>
            </tbody>
          </table>
        </blockquote>
        <p>
          <xsl:value-of select="documentation"/>
        </p>
      </body>
    </html>
  </xsl:template>
</xsl:stylesheet>