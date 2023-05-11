<?xml version="1.0"?>

<xsl:stylesheet version="1.0"
xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

  <xsl:template match="TestCaseMigratorPlus">
    <html>
      <body>
        <h2>Summary</h2>
        <table border="1" cellpadding="2">
          <tr>
            <td>
              Migrated successfully
            </td>
            <td>
              <a>
                <xsl:attribute name="href">
                  <xsl:value-of select="Summary/PassedWorkItems/@File"/>
                </xsl:attribute>
                <xsl:value-of select="Summary/PassedWorkItems/@File"/>
              </a>
            </td>
            <td>
              <xsl:value-of select="Summary/PassedWorkItems/@Count"/>
            </td>
          </tr>

          <tr>
            <td>
              Migrated with warning
            </td>
            <td>
              <a>
                <xsl:attribute name="href">
                  <xsl:value-of select="Summary/WarningWorkItems/@File"/>
                </xsl:attribute>
                <xsl:value-of select="Summary/WarningWorkItems/@File"/>
              </a>
            </td>
            <td>
              <xsl:value-of select="Summary/WarningWorkItems/@Count"/>
            </td>
          </tr>

          <tr>
            <td>
              Failed to migrate
            </td>
            <td>
              <a>
                <xsl:attribute name="href">
                  <xsl:value-of select="Summary/FailedWorkItems/@File"/>
                </xsl:attribute>
                <xsl:value-of select="Summary/FailedWorkItems/@File"/>
              </a>
            </td>
            <td>
              <xsl:value-of select="Summary/FailedWorkItems/@Count"/>
            </td>
          </tr>

        </table>

        <h3>Migrated successfully:</h3>
        <xsl:for-each select="WorkItems/PassedWorkItems/PassedWorkItem">
          Source:
          <a>
            <xsl:attribute name="href">
              <xsl:value-of select="@Source"/>
            </xsl:attribute>
            <xsl:value-of select="@Source"/>
          </a>
          <br />
          TFS Id:<xsl:value-of select="@TFSId"/>
          <br />
          <br />
        </xsl:for-each>

        <h3>Migrated with warning:</h3>
        <xsl:for-each select="WorkItems/WarningWorkItems/WarningWorkItem">
          Source:
          <a>
            <xsl:attribute name="href">
              <xsl:value-of select="@Source"/>
            </xsl:attribute>
            <xsl:value-of select="@Source"/>
          </a>
          <br />
          TFS Id:<xsl:value-of select="@TFSId"/>
          <br />
          Warning:<xsl:value-of select="@Warning"/>
          <br />
          <br />
        </xsl:for-each>

        <h3>Failed to migrate:</h3>
        <xsl:for-each select="WorkItems/FailedWorkItems/FailedWorkItem">
          Source:
          <a>
            <xsl:attribute name="href">
              <xsl:value-of select="@Source"/>
            </xsl:attribute>
            <xsl:value-of select="@Source"/>
          </a>
          <br />
          Error:<xsl:value-of select="@Error"/>
          <br />
          <br />
        </xsl:for-each>

        <br />
        <h4>Command Used:</h4>
        <xsl:value-of select="CommandLine"/>
      </body>
    </html>
  </xsl:template>

</xsl:stylesheet>