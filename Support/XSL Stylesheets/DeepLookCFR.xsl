<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
<xsl:output method="html"/>

   <!--
   ////////////////////////////////////////////////////////////////////////
                               Main Template
   ////////////////////////////////////////////////////////////////////////
   -->
   
   <xsl:template match="/">
   
      <html>

         <head>
            <meta http-equiv="Content-Type" content="text/html; charset=windows-1252"></meta>
            <title>DeepLook XML Report</title>
            <style>
               h1                   { color: SteelBlue; margin-left: 0 }
               h2                   { font-family: "Ariel", Arial; color: SteelBlue; font-size: 16; margin-left: 10}
               h3                   { font-family: "Times New Roman", Times New Roman; color: Black; font-size: 12; margin-left: 20 }
               p                    { font-family: "Times New Roman", Times New Roman; font-weight: Normal; color: Black; font-size: 10; margin-left: 35 }
               a                    { text-decoration: none; color: Blue }
               .PropertyName        { font-weight: bold }
               .ListHeader          { font-family: "Ariel", Arial; font-weight: bold; color: #000F80 }
               body                 { font-size: x-small }
               tbody                { font-size: x-small }
            </style>
         </head>

         <body bgcolor="BlanchedAlmond">
            <br />

            <xsl:apply-templates select="Project" />

	    <div align="center"><font face="Arial" color="#000080"><b>VB¼°.NET¹¤³ÌÔ´´úÂëÉ¨Ãè·ÖÎö¹¤¾ß V4.12.0</b></font></div>
            <div align="center"><font color="#000080">Please address all emails to &quot;dean_camera@hotmail.com&quot;</font></div>
	    <div align="center"><small><small>If this page fails to show correctly, update your webbrowser to a more recent version.</small></small></div>
         </body>

       </html>
   
   </xsl:template>
   
   
   <!--
   ////////////////////////////////////////////////////////////////////////
                               Template for the Project
   ////////////////////////////////////////////////////////////////////////
   -->
   
   <xsl:template match="Project">

           <h1>DeepLook Copy Required Files Report

           <xsl:if test="EXEName != '' ">
           (<xsl:value-of select="EXEName"/>)
           </xsl:if>
	   </h1>

	   <span class='PropertyName'>Report Generated: </span><xsl:value-of select="Created"/>


           <p class="Header1Paragraph">
              <xsl:apply-templates select="DLL" />
	      <br/>
              <xsl:apply-templates select="OCX" />
	      <br/>
              <xsl:apply-templates select="MISC" />
           </p>   
   </xsl:template>

   <!--
   ////////////////////////////////////////////////////////////////////////
                               Template for the DLLs
   ////////////////////////////////////////////////////////////////////////
   -->

   <xsl:template match="DLL">
        <h2>DLL Files</h2>

        <h3>The following files have been copied automatically:</h3>

         <xsl:for-each select="Copied">
            <p><xsl:value-of select="FileName" /></p>
         </xsl:for-each>	

        <h3>The following files may need to be copied manually:</h3>

         <xsl:for-each select="ManualCopy">
            <p><xsl:value-of select="FileName" /></p>
         </xsl:for-each>	

        <h3>The following files were not nessesary:</h3>

         <xsl:for-each select="NoCopy">
            <p><xsl:value-of select="FileName" /></p>
         </xsl:for-each>	
   </xsl:template>

   <!--
   ////////////////////////////////////////////////////////////////////////
                               Template for the OCXs
   ////////////////////////////////////////////////////////////////////////
   -->

   <xsl:template match="OCX">
        <h2>OCX Files</h2>

        <h3>The following files have been copied automatically:</h3>

         <xsl:for-each select="Copied">
            <p><xsl:value-of select="FileName" /></p>
         </xsl:for-each>	

        <h3>The following files may need to be copied manually:</h3>

         <xsl:for-each select="ManualCopy">
            <p><xsl:value-of select="FileName" /></p>
         </xsl:for-each>	

        <h3>The following files were not nessesary:</h3>

         <xsl:for-each select="NoCopy">
            <p><xsl:value-of select="FileName" /></p>
         </xsl:for-each>	
   </xsl:template>

   <!--
   ////////////////////////////////////////////////////////////////////////
                               Template for the MISC files
   ////////////////////////////////////////////////////////////////////////
   -->

   <xsl:template match="MISC">
        <h2>MISC Files</h2>

        <h3>The following files may need to be copied manually:</h3>

         <xsl:for-each select="ManualCopy">
            <p><xsl:value-of select="FileName" /></p>
         </xsl:for-each>	
   </xsl:template>

</xsl:stylesheet>
