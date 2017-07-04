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
               h2                   { color: SteelBlue; margin-left: 20 }
               h3                   { color: SteelBlue; margin-left: 40 }
               .Header1Paragraph    { margin-left: 20 }
               .Header2Paragraph    { margin-left: 40 }
               .Header3Paragraph    { margin-left: 60 }
               .Header4Paragraph    { margin-left: 80 }
               .Header5Paragraph    { margin-left: 100 }
               .Header6Paragraph    { margin-left: 120 }
               a                    { text-decoration: none; color: Blue }
               table                { border-collapse: collapse; margin-top: 10 }
               td                   { border: 1 solid #000000 }
               th                   { font-family: "Ariel", Arial; color: #FFFFFF; background-color: SteelBlue; text-align: left }
               .PropertyName        { font-weight: bold }
               .ListHeader          { font-family: "Ariel", Arial; font-weight: bold; color: #000F80 }
               body                 { font-size: x-small }
               tbody                { font-size: x-small }
               .r0                  { background-color: Bisque }
               .r1                  { background-color: BlanchedAlmond }               
            </style>
         </head>

         <body bgcolor="BlanchedAlmond">
            <br />

            <!-- This template applies to Project Group level documentation only -->
            <xsl:apply-templates select="ProjectGroup" />

            <!-- This template applies to Project level documentation only -->
            <xsl:apply-templates select="Project" />

            <!-- This template applies to File level documentation only -->
            <xsl:apply-templates select="File" />




	    <div align="center"><font face="Arial" color="#000080"><b>DeepLook - Freeware Visual Basic Project and File Scanner</b></font></div>
            <div align="center"><font color="#000080">Please address all emails to &quot;dean_camera@hotmail.com&quot;</font></div>
	    <div align="center"><small><small>If this page fails to show correctly, update your webbrowser to a more recent version.</small></small></div>
         </body>

       </html>
   
   </xsl:template>
   
   
   <!--
   ////////////////////////////////////////////////////////////////////////
                               Template for ProjectGroup
   ////////////////////////////////////////////////////////////////////////
   -->
   
   <xsl:template match="ProjectGroup">

           <h1>DeepLook Report

           <xsl:if test="FileName != '' ">
           (<xsl:value-of select="FileName"/>)
           </xsl:if>
	   </h1>

	   <span class='PropertyName'>Report Generated: </span><xsl:value-of select="Created"/>


      <p class="Header1Paragraph">
      
         <!-- Build the projects table -->
         <xsl:apply-templates select="Projects" />

      <br/>
      <hr/>

      </p>

      <!-- Build the documentation section for each project -->
      <xsl:for-each select="Projects/Project">

         <!-- Apply template for Project -->
         <xsl:apply-templates select="."/>

      </xsl:for-each>
   
   </xsl:template>

   <!--
   ////////////////////////////////////////////////////////////////////////
                               Template for Project
   ////////////////////////////////////////////////////////////////////////
   -->
   
   <xsl:template match="Project">

      <!-- Save the name of the project in a variable for later use -->
      <xsl:variable name="ProjectName"><xsl:value-of select="Name" /></xsl:variable>

      <!-- Build the header -->
      <h2>

         <!-- Build a bookmark -->
         <a>
            <xsl:attribute name="name">
               <xsl:call-template name="ProjectBookmarkName">
                  <xsl:with-param name="ProjectName"><xsl:value-of select="Name" /></xsl:with-param>
               </xsl:call-template>
            </xsl:attribute>
         </a>

         Project <xsl:value-of select="Name"/>

      </h2>
   
      <p class="Header2Paragraph">

         <!-- Build some properties of the project-->
         <span class='PropertyName'>Project Name: </span><xsl:value-of select="Name"/><br/>
         <span class='PropertyName'>Filename: </span><xsl:value-of select="FileName"/><br/>
         <span class='PropertyName'>Type: </span><xsl:value-of select="Type"/><br/>
         <span class='PropertyName'>Project Version: </span><xsl:value-of select="Description"/><br/>
         <br/>


      <span class='ListHeader'>Line Statistics:</span>

      <table border="0" width="400">

      <tr class="r0">
         <th>Statistic</th>
         <th>Value</th>
      </tr>
      <tr class="r1">
         <td width="200">Startup Item</td>
         <td width="200"><xsl:value-of select="LinStat1"/></td>
      </tr>
      <tr class="r0">
         <td width="200">Source Safe</td>
         <td width="200"><xsl:value-of select="LinStat2"/></td>
      </tr>
      <tr class="r1">
         <td width="200">Lines (Inc. Blanks)</td>
         <td width="200"><xsl:value-of select="LinStat3"/></td>
      </tr>
      <tr class="r0">
         <td width="200">Lines (No. Blanks)</td>
         <td width="200"><xsl:value-of select="LinStat4"/></td>
      </tr>
      <tr class="r1">
         <td width="200">Lines (Comments)</td>
         <td width="200"><xsl:value-of select="LinStat5"/></td>
      </tr>
      <tr class="r0">
         <td width="200">Declared Variables</td>
         <td width="200"><xsl:value-of select="LinStat6"/></td>
      </tr>

      </table>

      <br/>

      <span class='ListHeader'>File Count:</span>

      <table border="0" width="400">

      <tr>
         <th>Type</th>
         <th>Total</th>
      </tr>

      <tr class="r1">
         <td width="200">Forms:</td>
         <td width="200"><xsl:value-of select="FCount1"/></td>
      </tr>
      <tr class="r0">
         <td width="200">Modules:</td>
         <td width="200"><xsl:value-of select="FCount2"/></td>
      </tr>
      <tr class="r1">
         <td width="200">Class Modules:</td>
         <td width="200"><xsl:value-of select="FCount3"/></td>
      </tr>
      <tr class="r0">
         <td width="200">User Controls:</td>
         <td width="200"><xsl:value-of select="FCount4"/></td>
      </tr>
      <tr class="r1">
         <td width="200">User Documents:</td>
         <td width="200"><xsl:value-of select="FCount5"/></td>
      </tr>
      <tr class="r0">
         <td width="200">Property Pages:</td>
         <td width="200"><xsl:value-of select="FCount6"/></td>
      </tr>
      <tr class="r1">
         <td width="200">Designers:</td>
         <td width="200"><xsl:value-of select="FCount7"/></td>
      </tr>

      </table>

      <br/>

      <span class='ListHeader'>Sub/Function/Property Statistics:</span>

      <table border="0" width="400">

      <tr>
         <th>Statistic</th>
         <th>Value</th>
      </tr>

      <tr class="r1">
         <td width="200">Subs:</td>
         <td width="200"><xsl:value-of select="SPFInfo1"/></td>
      </tr>
      <tr class="r0">
         <td width="200">Functions:</td>
         <td width="200"><xsl:value-of select="SPFInfo2"/></td>
      </tr>
      <tr class="r1">
         <td width="200">Properties:</td>
         <td width="200"><xsl:value-of select="SPFInfo3"/></td>
      </tr>
      <tr class="r0">
         <td width="200">Events:</td>
         <td width="200"><xsl:value-of select="SPFInfo4"/></td>
      </tr>
      <tr class="r1">
         <td width="200">Declared External Subs:</td>
         <td width="200"><xsl:value-of select="SPFInfo5"/></td>
      </tr>
      <tr class="r0">
         <td width="200">Declared External Functions:</td>
         <td width="200"><xsl:value-of select="SPFInfo6"/></td>
      </tr>

      </table>

      <br/>

         <!-- Build the references table -->
         <xsl:apply-templates select="References" />

         <br/>

         <!-- Build the declared DLLs table -->
         <xsl:apply-templates select="DeclaredDLLs" />

         <br/>      

         <!-- Build the related documents table -->
         <xsl:apply-templates select="RelDocs" />

         <br/> 

         <!-- Build the files table -->
         <xsl:apply-templates select="Files" >
            <xsl:with-param name="ProjectName"><xsl:value-of select="Name" /></xsl:with-param>
         </xsl:apply-templates>

         <br/>
      
         <!-- Build the documentation section for each file -->
         <xsl:for-each select="Files/File">
            <xsl:sort select="Type"/>
            <xsl:sort select="Name"/>
            <!-- Apply template for File -->
            <xsl:apply-templates select=".">
               <xsl:with-param name="ProjectName"><xsl:value-of select="$ProjectName" /></xsl:with-param>
            </xsl:apply-templates>
     
         </xsl:for-each>

      </p>
   
      <br/>
      <hr/>
      <br/>
   
   </xsl:template>

   <!--
   ////////////////////////////////////////////////////////////////////////
                               Template for File
   ////////////////////////////////////////////////////////////////////////
   -->
   
   <xsl:template match="File">
      <xsl:param name="ProjectName"></xsl:param>   

      <!-- Save the name of the file in a variable for later use -->
      <xsl:variable name="FileName"><xsl:value-of select="Name" /></xsl:variable>
   
         <h3>

            <!-- Build a bookmark name -->
            <a>
               <xsl:attribute name="name">
                  <xsl:call-template name="FileBookmarkName">
                     <xsl:with-param name="ProjectName"><xsl:value-of select="$ProjectName" /></xsl:with-param>
                     <xsl:with-param name="FileName"><xsl:value-of select="Name" /></xsl:with-param>
                  </xsl:call-template>
               </xsl:attribute>
            </a>

            <!-- Add the translated file type -->
            <xsl:apply-templates select="Type" />

            <!-- Add a space character as separator -->
            <xsl:text> </xsl:text>

            <!-- Add the logical name and the file name -->
            <xsl:value-of select="Name"/>
            <xsl:if test="FileName != '' ">
               <xsl:text> (</xsl:text>
               <xsl:value-of select="FileName"/>
               <xsl:text>)</xsl:text>
            </xsl:if>

         </h3>

         <p class="Header3Paragraph">

            <span class='PropertyName'>FileName: </span><xsl:value-of select="FileName"/><br/>
            <span class='PropertyName'>Name: </span><xsl:value-of select="Name"/><br/>
            <br/>
            <span class='PropertyName'>Lines (Inc. Blanks): </span><xsl:value-of select="FStat1"/><br/>
            <span class='PropertyName'>Lines (No. Blanks): </span><xsl:value-of select="FStat2"/><br/>
            <span class='PropertyName'>Lines (Comment): </span><xsl:value-of select="FStat3"/><br/>
            <span class='PropertyName'>Lines (Hybrid): </span><xsl:value-of select="FStat4"/><br/>
            <br/>

	    <xsl:if test="Type = 'Form' or Type = 'User Control' or Type = 'User Document' or Type = 'Designer' or Type = 'Property Page' ">
            <span class='PropertyName'>Controls: </span><xsl:value-of select="FStat5"/><br/>
            <span class='PropertyName'>Variables: </span><xsl:value-of select="FStat6"/><br/>
            <br/>
            <span class='PropertyName'>Subroutines: </span><xsl:value-of select="FStat7"/><br/>
            <span class='PropertyName'>Functions: </span><xsl:value-of select="FStat8"/><br/>
            <span class='PropertyName'>Properties: </span><xsl:value-of select="FStat9"/><br/>
            <span class='PropertyName'>Events: </span><xsl:value-of select="FStat10"/><br/>
	    </xsl:if>

	    <xsl:if test="Type = 'Class Module' or Type = 'Module'">
            <span class='PropertyName'>Variables: </span><xsl:value-of select="FStat5"/><br/>
            <br/>
            <span class='PropertyName'>Subroutines: </span><xsl:value-of select="FStat6"/><br/>
            <span class='PropertyName'>Functions: </span><xsl:value-of select="FStat7"/><br/>
            <span class='PropertyName'>Properties: </span><xsl:value-of select="FStat8"/><br/>

	    <xsl:if test="Type = 'Class Module' ">
            <span class='PropertyName'>Events: </span><xsl:value-of select="FStat9"/><br/>
	    </xsl:if>

	    </xsl:if>

            <br/>
         </p>
   
   </xsl:template>

   <!--
   ////////////////////////////////////////////////////////////////////////
                               Template for Project list
   ////////////////////////////////////////////////////////////////////////
   -->

   <xsl:template match="Projects">

      <span class='ListHeader'>Projects:</span>

      <table border="0" width="90%">

      <tr>
         <th>Project Name</th>
         <th>Filename</th>
         <th>Project Type</th>
      </tr>

      <xsl:for-each select="Project">

         <xsl:sort select="Type"/>
         <xsl:sort select="Name"/>
       
         <tr class="r{position() mod 2}">
            <td>
               <!-- We add a hyperlink to a bookmark with information about of the project -->
               <a>
                  <xsl:attribute name="href">
                     <xsl:text>#</xsl:text>
                     <xsl:call-template name="ProjectBookmarkName">
                        <xsl:with-param name="ProjectName"><xsl:value-of select="Name" /></xsl:with-param>
                     </xsl:call-template>
                  </xsl:attribute>
                  <xsl:value-of select="Name" />
               </a>
            </td>
            <td><xsl:value-of select="FileName"/></td>
            <td><xsl:value-of select="Type"/></td>
         </tr>

       </xsl:for-each>

      </table>

   </xsl:template>

   <!--
   ////////////////////////////////////////////////////////////////////////
                               Template for Files list
   ////////////////////////////////////////////////////////////////////////
   -->

   <xsl:template match="Files">

      <xsl:param name="ProjectName"></xsl:param>   

      <span class='ListHeader'>Files:</span>

      <table border="0" width="90%">

         <tr>
            <th>Name</th>
            <th>Filename</th>
            <th>Type</th>
         </tr>

         <xsl:for-each select="File">
            <xsl:sort select="Type" />
            <xsl:sort select="Name" />
            <tr class="r{position() mod 2}">
               <td>
                  <!-- We add a hyperlink to a bookmark with information about of the file -->
                  <a>
                     <xsl:attribute name="href">
                        <xsl:text>#</xsl:text>
                        <xsl:call-template name="FileBookmarkName">
                           <xsl:with-param name="ProjectName"><xsl:value-of select="$ProjectName" /></xsl:with-param>
                           <xsl:with-param name="FileName"><xsl:value-of select="Name" /></xsl:with-param>
                        </xsl:call-template>
                     </xsl:attribute>
                     <xsl:value-of select="Name" />
                  </a>
               </td>
               <td><xsl:value-of select="FileName"/></td>
               <td>
                  <xsl:apply-templates select="Type"/>
               </td>
            </tr>
         </xsl:for-each>

      </table>

   </xsl:template>

   <!--
   ////////////////////////////////////////////////////////////////////////
                               Template for References list
   ////////////////////////////////////////////////////////////////////////
   -->

   <xsl:template match="References">

      <span class='ListHeader'>References:</span>

      <table border="0" width="400">

      <tr>
         <th>Filename</th>
         <th>Description</th>
      </tr>
          
      <xsl:for-each select="Reference">

         <tr class="r{position() mod 2}">
            <td><xsl:value-of select="FileName"/></td>
            <td><xsl:value-of select="Description"/></td>
         </tr>

      </xsl:for-each>

      </table>

   </xsl:template>

   <!--
   ////////////////////////////////////////////////////////////////////////
                           Template for Declared DLLs list
   ////////////////////////////////////////////////////////////////////////
   -->

   <xsl:template match="DeclaredDLLs">

      <span class='ListHeader'>Declared DLL Files:</span>

      <table border="0" width="400">

      <tr>
         <th>Filename</th>
      </tr>
          
      <xsl:for-each select="DeclaredDLL">

         <tr class="r{position() mod 2}">
            <td><xsl:value-of select="FileName"/></td>
         </tr>

      </xsl:for-each>

      </table>

   </xsl:template>

   <!--
   ////////////////////////////////////////////////////////////////////////
                       Template for Related Documents list
   ////////////////////////////////////////////////////////////////////////
   -->

   <xsl:template match="RelDocs">

      <span class='ListHeader'>Related Documents:</span>

      <table border="0" width="400">

      <tr>
         <th>Filename</th>
      </tr>
          
      <xsl:for-each select="RelDoc">

         <tr class="r{position() mod 2}">
            <td><xsl:value-of select="FileName"/></td>
         </tr>

      </xsl:for-each>

      </table>

   </xsl:template>

   <!--
   ////////////////////////////////////////////////////////////////////////
                               Template for FileType
   ////////////////////////////////////////////////////////////////////////
   -->
   
   <xsl:template match="File/Type">

      <xsl:choose>

         <!-- Here you can localize the type to your language -->

         <xsl:when test=". = 'Module' ">Module</xsl:when>
         <xsl:when test=". = 'Class Module' ">Class</xsl:when>
         <xsl:when test=". = 'Form' ">Form</xsl:when>
         <xsl:when test=". = 'Property Page' ">Property Page</xsl:when>
         <xsl:when test=". = 'User Control' ">User Control</xsl:when>
         <xsl:when test=". = 'User Document' ">User Document</xsl:when>
         <xsl:when test=". = 'Designer' ">ActiveX Designer</xsl:when>

      </xsl:choose>

   </xsl:template>
   
   <!--
   ////////////////////////////////////////////////////////////////////////
                               Template for Project Bookmark
   ////////////////////////////////////////////////////////////////////////
   -->
   
   <xsl:template name="ProjectBookmarkName">
      <xsl:param name="ProjectName"></xsl:param>   

      <!-- Build a bookmark name with the pattern "ProjectName_Bookmark" -->
      <xsl:value-of select="$ProjectName" />
      <xsl:text>_Bookmark</xsl:text>

   </xsl:template>
 
   <!--
   ////////////////////////////////////////////////////////////////////////
                               Template for File Bookmark
   ////////////////////////////////////////////////////////////////////////
   -->
   
   <xsl:template name="FileBookmarkName">

      <xsl:param name="ProjectName"></xsl:param>   
      <xsl:param name="FileName"></xsl:param>   

      <!-- Build a bookmark name with the pattern "ProjectName_FileName_Bookmark" -->
      <xsl:value-of select="$ProjectName" />
      <xsl:text>_</xsl:text>
      <xsl:value-of select="$FileName"/>
      <xsl:text>_Bookmark</xsl:text>

   </xsl:template>

</xsl:stylesheet>
