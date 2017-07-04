Attribute VB_Name = "ModXMLReport"
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:11:54
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：ModXMLReport
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:11:54
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V4.12.0
'*************************************************************************
'-----------------------------------------------------------------------------------------------
'                                    XML REPORT GENERATOR SYNOPSIS
'-----------------------------------------------------------------------------------------------
'   Yes, this is yet another parser engine. It will accept a standard DeepLook VB6 text report
'   as its input, generating a XML report file as its output. It is very fast and uses an idea
'   I got from my microcontroller (programmable electronic ICs) exploits; the use of a stack to
'   make the generation easier, as every XML header must be terminated. Paired with ClsXMLStack.
'
'   This module also has the routines for generating the copy report XML files.
'-----------------------------------------------------------------------------------------------

Option Explicit

'-----------------------------------------------------------------------------------------------
Private cStack As ClsXMLStack
Private TXTReportFileNum As Integer
'-----------------------------------------------------------------------------------------------

Public Sub MakeXMLReport(FileName As String, ReportFileName As String)
    Dim Buffer As String

    Set cStack = New ClsXMLStack ' Create new instance of the XML stack class and open a new XML file
    cStack.OpenFile FileName ' Open the XML file

    ' Add start XML template info:
    cStack.Add "<?xml version=""1.0"" encoding=""Windows-1252"" ?>"
    cStack.Add "<?xml-stylesheet type=""text/xsl"" href=""DeepLook.xsl"" ?>"
    cStack.Push "<ProjectGroup>"

    cStack.PushAddPop "<Created>", Now

    TXTReportFileNum = FreeFile ' Get a free file number
    Open ReportFileName For Input As TXTReportFileNum ' Open the created temp text report

    If SkipLinesAndGet(2) <> "===================DEEPLOOK PROJECT REPORT===================" Then
        MsgBoxEx "Not a valid DeepLook text report file!", vbCritical, "DeepLook - XML Report Error", , , , , PicError
        Set cStack = Nothing
        Kill FileName ' Delete the empty XML file
    Else
        Do While Not EOF(TXTReportFileNum)
            Line Input #TXTReportFileNum, Buffer ' Read in a line of the text report
            ParseTXTReportLine Buffer ' Parse the report line
        Loop

        Close TXTReportFileNum ' Close the tempary text report file

        If cStack.StackDepth Then cStack.PopAll ' Stack is not empty; Pop remaining headers off the XML stack
        Set cStack = Nothing ' Delete the instace of the XML stack (XML file automatically closed)
    End If
End Sub

Public Sub MakeXMLCopyReport(FileName As String)
    Dim XMLFileNum As Long

    FrmResults.sbrStatus.Caption = "STAT>正在创建 XML 报告副本..."

    Set cStack = New ClsXMLStack ' Create new instance of the XML stack class and open a new XML file
    cStack.OpenFile FileName ' Open the XML file

    ' Add start XML template info:
    cStack.Add "<?xml version=""1.0"" encoding=""Windows-1252"" ?>"
    cStack.Add "<?xml-stylesheet type=""text/xsl"" href=""DeepLookCFR.xsl"" ?>"

    cStack.Push "<Project>"

    cStack.PushAddPop "<Created>", Now

    On Local Error Resume Next
    If FrmResults.TreeView.Nodes("PROJECT_?_EXESTATS_NAME") Then ' If node not found, the XML item shouldn't be added at all
        cStack.PushAddPop "<EXEName>", Mid$(FrmResults.TreeView.Nodes("PROJECT_?_EXESTATS_NAME"), 12)
    End If
    On Local Error GoTo 0

    AddCopyReportItems "DLL" ' Add DLL files
    AddCopyReportItems "OCX" ' Add OCX files
    AddCopyReportMiscFiles ' Add MISC files

    cStack.Pop ' Pop <Project>

    If cStack.StackDepth Then cStack.PopAll ' Stack is not empty; Pop remaining headers off the XML stack
    Set cStack = Nothing ' Delete the instace of the XML stack (XML file automatically closed)

    XMLFileNum = FreeFile
    Open GetRootDirectory(FileName) & "DeepLookCFR.xsl" For Output As #XMLFileNum ' Open a blank XML stylesheet

    Dim StrData As String
    StrData = StrConv(LoadResData(2, "XMLTEMPLATE"), vbUnicode)
    StrData = IIf(Right(StrData, 2) <> "t>", Left$(StrData, Len(StrData) - 2), StrData)
    Print #XMLFileNum, StrData ' Write the XML stylesheet into the created file

    Close #XMLFileNum ' Close the XML Stylesheet file
End Sub

Public Sub FinishXMLCopyReport(FileName As String)
    Dim Buffer As String
    Dim Report As String
    Dim XMLFileNum As Long
    Dim NodeNum As Long

    XMLFileNum = FreeFile
    Open FileName For Input As #XMLFileNum
    Do Until EOF(XMLFileNum)
        Line Input #XMLFileNum, Buffer
        Report = Report & IIf(Report = vbNullString, vbNullString, vbCrLf) & Buffer
    Loop
    Close #XMLFileNum

    With FrmCopyReport.tvwItemsTV
        For NodeNum = 1 To .Nodes.Count
            If .Nodes(NodeNum).Image = "Error" Then
                Report = Replace(Report, .Nodes(NodeNum).Text, .Nodes(NodeNum).Text & " [Could not Copy]", 1, 1)
            End If
        Next
    End With

    XMLFileNum = FreeFile
    Open FileName For Output As #XMLFileNum
    Print #1, Report
    Close #XMLFileNum
End Sub

Private Sub ParseTXTReportLine(linedata As String)
    Dim I As Integer

    If Left$(Trim$(linedata), 13) = "VISUAL BASIC " Then
        If cStack.Peek <> "<Files>" Then cStack.Push "<Files>"

        cStack.Push "<File>"

        cStack.PushAddPop "<Type>", StrConv(Mid$(Trim$(linedata), 14), vbProperCase)

        linedata = SkipLinesAndGet(2)
        cStack.PushAddPop "<FileName>", Mid$(linedata, 27)
        Line Input #TXTReportFileNum, linedata
        cStack.PushAddPop "<Name>", Mid$(linedata, 27)

        linedata = SkipLinesAndGet(2)

        Do
            If LenB(linedata) Then
                I = I + 1
                cStack.PushAddPop "<FStat" & I & ">", Mid$(linedata, 27)
            End If

            Line Input #TXTReportFileNum, linedata
        Loop While Not EOF(TXTReportFileNum) And InStr(1, linedata, "===") = 0

        cStack.Pop ' Pop <File>
    ElseIf InStrB(1, linedata, "Visual Basic Project File") Then
        On Local Error Resume Next

        If cStack.Peek = "<Files>" Then
            cStack.Pop ' Pop <Files>
            cStack.Pop ' Pop <Project>
        ElseIf cStack.Peek = "<Project>" Then
            cStack.Pop  ' Pop <Project>
        ElseIf cStack.Peek = "<ProjectGroup>" Then
            cStack.Push "<Projects>"
        End If

        linedata = SkipLinesAndGet(2)

        cStack.Push "<Project>"
        cStack.PushAddPop "<FileName>", Mid$(linedata, 27)

        Line Input #TXTReportFileNum, linedata
        cStack.PushAddPop "<Name>", Mid$(linedata, 27)

        Line Input #TXTReportFileNum, linedata
        cStack.PushAddPop "<Description>", Mid$(linedata, 27)

        linedata = SkipLinesAndGet(2)
        cStack.PushAddPop "<LinStat1>", Mid$(linedata, 27)
        Line Input #TXTReportFileNum, linedata
        cStack.PushAddPop "<LinStat2>", Mid$(linedata, 27)
        linedata = SkipLinesAndGet(2)
        cStack.PushAddPop "<Type>", Mid$(linedata, 27)

        For I = 3 To 6
            Line Input #TXTReportFileNum, linedata
            cStack.PushAddPop "<LinStat" & I & ">", Format$(Int(Mid$(linedata, 27)), "###,###,###")
        Next

        Line Input #TXTReportFileNum, linedata

        For I = 1 To 7
            Line Input #TXTReportFileNum, linedata
            cStack.PushAddPop "<FCount" & I & ">", Mid$(linedata, 27)
        Next

        linedata = SkipLinesAndGet(3)

        For I = 1 To 6
            Line Input #TXTReportFileNum, linedata
            cStack.PushAddPop "<SPFInfo" & I & ">", Mid$(linedata, 27)
        Next

        cStack.Push "<References>"

        Do
            Line Input #TXTReportFileNum, linedata ' Ensure that we are in the correct place for the tables
        Loop Until InStrB(1, linedata, "---------")
            
        linedata = SkipLinesAndGet(2)

        Do
            cStack.Push "<Reference>"

            cStack.PushAddPop "<FileName>", Trim$(Left$(linedata, InStr(1, linedata, "|") - 2))
            linedata = Mid$(linedata, InStr(1, linedata, "|") + 2)
            cStack.PushAddPop "<Description>", Trim$(Left$(linedata, InStr(1, linedata, "|") - 2))

            cStack.Pop ' Pop <Reference>

            Line Input #TXTReportFileNum, linedata
        Loop While InStrB(1, linedata, "-+-") = 0 And LenB(linedata)

        cStack.Pop ' Pop <References>

        linedata = SkipLinesAndGet(7)
        cStack.Push "<DeclaredDLLs>"

        Do
            cStack.Push "<DeclaredDLL>"
            cStack.PushAddPop "<FileName>", Trim$(Left$(linedata, Len(linedata) - 1))
            cStack.Pop ' Pop <DeclaredDLL>
            Line Input #TXTReportFileNum, linedata
        Loop While InStrB(1, linedata, "-----+") = 0 And LenB(linedata)

        cStack.Pop ' Pop <DeclaredDLLs>
        linedata = SkipLinesAndGet(7)
        cStack.Push "<RelDocs>"

        Do
            cStack.Push "<RelDoc>"
            cStack.PushAddPop "<FileName>", Trim$(Left$(linedata, Len(linedata) - 1))
            cStack.Pop ' Pop <RelDoc>
            Line Input #TXTReportFileNum, linedata
        Loop While InStrB(1, linedata, "-----+") = 0 And LenB(linedata)

        cStack.Pop ' Pop <RelDocs>
    ElseIf InStrB(1, linedata, "**Visual Basic Group**") Then
        linedata = SkipLinesAndGet(2)

        cStack.PushAddPop "<FileName>", Mid$(linedata, 31)
    End If
End Sub

Private Function SkipLinesAndGet(Count As Integer) As String
    Dim I As Integer

    For I = 1 To Count
        Line Input #TXTReportFileNum, SkipLinesAndGet
    Next
End Function

Public Sub ConvertTXTtoXML(TXTFileName As String)
    Dim RetVal As Integer, I As Integer

    With FrmSelProject.lblScanPhase ' Only visible when converting TXT report to XML from the select project form
        .ForeColor = RGB(50, 50, 180)
        .Caption = "Ready for XML Conversion..."
        .Visible = True
    End With

    With FrmSelProject.cdgDialogs
        .Filter = "DeepLook XML Report File (*.xml)|*.xml"
        .FileName = vbNullString
        .ShowSave

        If .FileName = vbNullString Then
            FrmSelProject.lblScanPhase.Caption = "XML Conversion Canceled."
            Exit Sub
        End If

        If FileExists(.FileName) Then
            RetVal = MsgBoxEx("File """ & .FileName & """ already exists. Overwrite?", vbExclamation Or vbYesNo Or vbDefaultButton2, "DeepLook", , , , , PicReport)
            If RetVal = vbNo Then Exit Sub
        End If

        FrmSelProject.lblScanPhase.Caption = "Converting TXT Report to XML..."
        ModXMLReport.MakeXMLReport .FileName, TXTFileName ' Convert the XML File

        RetVal = FreeFile
        Open GetRootDirectory(.FileName) & "DeepLook.xsl" For Binary As #RetVal
        Put #RetVal, , LoadResData(1, "XMLTEMPLATE") ' Save the XML Style Template from the RES file

        For I = 1 To 12 ' Required, deletes the 12 bytes of junk that the RES file adds to the start of every LoadResData command
            Put #RetVal, I, " "
        Next

        Close #RetVal

        FrmSelProject.lblScanPhase.Caption = "XML 文件转换完毕."
    End With
End Sub

Private Sub AddCopyReportItems(FileType As String)
    Dim NodeIndex As Long
    Dim FilesPresent As Boolean

    cStack.Push "<" & FileType & ">"

    FilesPresent = False ' Flag, true if at least one of the type of file is present
    With FrmCopyReport.tvwItemsTV.Nodes
        For NodeIndex = 1 To .Count
            If .Item(NodeIndex).Checked Then
            If InStr(1, UCase$(.Item(NodeIndex).Text), "." & FileType) Then
                cStack.Push "<Copied>"
                cStack.PushAddPop "<FileName>", .Item(NodeIndex).Text
                cStack.Pop ' Pop <Copied>
                FilesPresent = True ' At least one of this type of file present, set the flag
            End If
            End If
        Next
    End With

    If Not FilesPresent Then
        cStack.Push "<Copied>"
        cStack.PushAddPop "<FileName>", "(无)"
        cStack.Pop ' Pop <Copied>
    End If

    FilesPresent = False ' Flag, true if at least one of the type of file is present
    With FrmCopyReport.tvwManualCopyTV.Nodes
        For NodeIndex = 1 To .Count
            If InStr(1, UCase$(.Item(NodeIndex).Text), "." & FileType) Then
                cStack.Push "<ManualCopy>"
                cStack.PushAddPop "<FileName>", .Item(NodeIndex).Text
                cStack.Pop ' Pop <ManualCopy>
                FilesPresent = True ' At least one of this type of file present, set the flag
            End If
        Next
    End With

    If Not FilesPresent Then
        cStack.Push "<ManualCopy>"
        cStack.PushAddPop "<FileName>", "(None)"
        cStack.Pop ' Pop <ManualCopy>
    End If

    FilesPresent = False ' Flag, true if at least one of the type of file is present
    With FrmCopyReport.tvwNonCopyItemsTV.Nodes
        For NodeIndex = 1 To .Count
            If InStr(1, UCase$(.Item(NodeIndex).Text), "." & FileType) Then
                cStack.Push "<NoCopy>"
                cStack.PushAddPop "<FileName>", .Item(NodeIndex).Text
                cStack.Pop ' Pop <NoCopy>
                FilesPresent = True ' At least one of this type of file present, set the flag
            End If
        Next
    End With

    If Not FilesPresent Then
        cStack.Push "<NoCopy>"
        cStack.PushAddPop "<FileName>", "(None)"
        cStack.Pop ' Pop <NoCopy>
    End If

    cStack.Pop ' Pop <#FileType#>
End Sub

Private Sub AddCopyReportMiscFiles()
    Dim NodeIndex As Long
    Dim FilesPresent As Boolean

    cStack.Push "<MISC>"

    FilesPresent = False ' Flag, true if at least one of the type of file is present
    With FrmCopyReport.tvwManualCopyTV.Nodes
        For NodeIndex = 1 To .Count
            If InStr(1, .Item(NodeIndex).Text, ".") = 0 Then
                cStack.Push "<ManualCopy>"
                cStack.PushAddPop "<FileName>", .Item(NodeIndex).Text
                cStack.Pop ' Pop <ManualCopy>
                FilesPresent = True ' At least one of this type of file present, set the flag
            End If
        Next
    End With

    If Not FilesPresent Then
        cStack.Push "<ManualCopy>"
        cStack.PushAddPop "<FileName>", "(None)"
        cStack.Pop ' Pop <ManualCopy>
    End If

    cStack.Pop ' Pop <MISC>
End Sub
