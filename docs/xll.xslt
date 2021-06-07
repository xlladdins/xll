<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:output method="html"
              encoding="utf-8"
              version="4.0"
              doctype-public="-//W3C//DTD HTML 4.01//EN"
              omit-xml-declaration="yes"
              indent="yes"
  />
  <xsl:template match="/args">
    <html lang="en">
      <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
        <link href="xll.css" rel="stylesheet" type="text/css" />
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.13.11/dist/katex.min.css" integrity="sha384-Um5gpz1odJg5Z4HAmzPtgZKdTBHZdw8S29IecapCSB31ligYPhHQZMIlWLYQGVoc" crossorigin="anonymous"/>
        <xsl:element name="script">
          <xsl:attribute name="defer"></xsl:attribute>
          <xsl:attribute name="src">https://cdn.jsdelivr.net/npm/katex@0.13.11/dist/katex.min.js</xsl:attribute>
          <xsl:attribute name="integrity">sha384-YNHdsYkH6gMx9y3mRkmcJ2mFUjTd0qNQQvY9VYZgQd7DcN7env35GzlmFaZ23JGp</xsl:attribute>
          <xsl:attribute name="crossorigin">anonymous</xsl:attribute>
        </xsl:element>
        <xsl:element name="script">
          <xsl:attribute name="defer"></xsl:attribute>
          <xsl:attribute name="src">https://cdn.jsdelivr.net/npm/katex@0.13.11/dist/contrib/auto-render.min.js</xsl:attribute>
          <xsl:attribute name="integrity">sha384-vZTG03m+2yp6N6BNi5iM4rW4oIwk5DfcNdFfxkk9ZWpDriOkXX8voJBFrAO7MpVl</xsl:attribute>
          <xsl:attribute name="crossorigin">anonymous</xsl:attribute>
          <xsl:attribute name="onload">renderMathInElement(document.body, {fleqn: true});</xsl:attribute>
        </xsl:element>
        <title>
          <xsl:value-of select="title"/>
        </title>
      </head>
      <body>
        <h1>
          <xsl:value-of select="functionText"/>
          <xsl:text> </xsl:text>
          <xsl:value-of select="type"/>
        </h1>
        <p>
          This article describes the formula syntax of the
          <xsl:value-of select="functionText"/>
          <xsl:text> </xsl:text>
          <xsl:value-of select="type"/>.
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
        <footer>
          Return to <a href="index.html">index</a>.
        </footer>
      </body>
    </html>
  </xsl:template>
</xsl:stylesheet>