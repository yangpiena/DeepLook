Attribute VB_Name = "ModAnalyseDOTNET"
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:11:00
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：ModAnalyseDOTNET
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:11:00
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V4.12.0
'*************************************************************************
'-----------------------------------------------------------------------------------------------
'                                  .NET SCANNING SYNOPSIS
'-----------------------------------------------------------------------------------------------
'   DeepLook is an advanced VB scanner. This module will scan some aspects of a .NET project
'   and display them on a treeview control, but it is still in its infancy and as such can only
'   give basic line stats and references. The .NET scanning engine does offer basic text reports.
'
'   The .NET scanning engine contained in this module is (C) Dean Camera.
'-----------------------------------------------------------------------------------------------

Option Explicit

' -----------------------------------------------------------------------------------------------
Private vbFile As ClsNETvbFile
Private vbProject As ClsNETproject

Private ScanningVBfile As Boolean
Private CodeDone As Boolean

Private ControlsInFile As Long
Private WaitForRegion As Boolean

Private FileType As String
Private ShowFNFerrors As Integer

Private ProjFileNum As Integer, VBFileNum As Integer
' -----------------------------------------------------------------------------------------------

Private Sub PrepareForScan() ' Clears report, resises controls, etc. ready for scan
    ShowCurrItemPic = GetSetting("DeepLook", "Options", "ShowCurrItemPic", 1) ' /
    ShowFNFerrors = GetSetting("DeepLook", "Options", "ShowFNFErrors", 1)

    If ShowCurrItemPic = 1 Then                             '  \
        FrmSelProject.pgbAPB.Width = 4215                   '  |
        FrmSelProject.imgCurrScanObjType.Visible = True     '  |  Rearrange items on the SelProject form
    Else                                                    '  |  depending on the settings retrieved
        FrmSelProject.pgbAPB.Width = 4695                   '  |
        FrmSelProject.imgCurrScanObjType.Visible = False    '  |
    End If                                                  '  /

    FrmReport.rtbReportText.Text = vbNullString ' Clear report
End Sub

Private Function ExtractFilePath(FileNameAndPath As String) As String ' Gets only the path from a file & path string
    Dim SlashPos As String

    SlashPos = InStrRev(FileNameAndPath, "\") ' Retrieve the position of the last path slash

    ExtractFilePath = FileNameAndPath

    If SlashPos = 0 Then Exit Function

    ExtractFilePath = Left$(FileNameAndPath, SlashPos) ' Trim path
End Function

Public Sub AnalyseDotNetProject(Path As String) ' Main sub to scan a .NET project
    Dim linedata As String, ReportPrjData As String, ReportPrjData2 As String
    Dim DonePrj As Long, LoopVar As Long

    IsScanning = True ' Set the scanning flag to true

    CurrentScanFile "NETProject" ' Change current scan icon to a VB.NET project

    PrepareForScan ' Resize controls, clear report, etc.
    AddReportHeader ' Add the Header info to the report

    PROJECTMODE = NET ' Set project mode to .NET

    If Not FileExists(Path) Then  ' Make sure the file exists
        MsgBoxEx "工程文件没有找到!", vbCritical, "扫描出错", , , , , PicError, , 5
        IsScanning = False ' Set the scanning flag to false
        Exit Sub
    End If

    ProjFileNum = FreeFile
    Open Path For Input As #ProjFileNum

    FrmSelProject.lblScanningName.Caption = ExtractFileName(Path)
    Set vbProject = New ClsNETproject ' Create a new instance to hold project stats

    Line Input #ProjFileNum, linedata ' Check for valid .NET Project Header
    If Trim$(linedata) <> "<VisualStudioProject>" Then GoTo NotVBNetProj
    Line Input #ProjFileNum, linedata ' Check for valid VB.NET Project Header
    If Trim$(linedata) <> "<VisualBasic" Then GoTo NotVBNetProj

    AddReportText vbNewLine & "============================================================="
    AddReportText "================Visual Basic .NET Project File==============="
    AddReportText "============================================================="
    AddReportText "               File Name: " & ExtractFileName(Path)
    AddReportText "            Project Name: " & "?PLACE>ProjectName"
    AddReportText "         Project Version: " & "?PLACE>ProjectVersion"
    AddReportText "============================================================="
    AddReportText "?PLACE>ProjectStats"

    Do While Not EOF(ProjFileNum)
        FrmSelProject.pgbAPB.Value = (95 / LOF(ProjFileNum)) * DonePrj ' Increase progressbar
        Line Input #ProjFileNum, linedata
        LookAtDOTNETProjLine linedata ' Scan the current line
        DonePrj = DonePrj + Len(linedata)
        If GetInputState Then DoEvents ' Process system events if there is messages in the keyboard/mouse buffers
    Loop

    Close #ProjFileNum

    With FrmResults.TreeView.Nodes ' Add the statistic nodes
        .Item(GetNodeNum("NETLines")).Text = "Lines (Inc. Blanks): " & vbProject.CodeLines
        .Item(GetNodeNum("NETLinesNB")).Text = "Lines (No Blanks): " & vbProject.CodeLinesNB
        .Item(GetNodeNum("NETLinesB")).Text = "Lines (Blanks): " & vbProject.BlankLines
        .Item(GetNodeNum("NETLinesComment")).Text = "Lines (Comment): " & vbProject.CommentLines
    End With

    TotalLines = vbProject.CodeLines + vbProject.CommentLines  ' Set the variable for layer calculation of lines/sec

    FrmReport.rtbReportText.Text = Replace$(FrmReport.rtbReportText.Text, "?PLACE>ProjectName", Mid$(FrmResults.TreeView.Nodes(GetNodeNum("NET_ProjectInfo_Name")).Text, 15)) ' Replace temp name in report with project name
    FrmReport.rtbReportText.Text = Replace$(FrmReport.rtbReportText.Text, "?PLACE>ProjectVersion", Mid$(FrmResults.TreeView.Nodes(GetNodeNum("NET_AssembleInfo_Version")).Text, 10)) ' Replace temp name in report with project version

    ReportPrjData = vbNewLine & "Lines (Inc. Blanks): " & vbProject.CodeLines & _
        vbNewLine & "Lines (No. Blanks): " & vbProject.CodeLinesNB & _
        vbNewLine & "Lines (Blanks): " & vbProject.BlankLines & _
        vbNewLine & "Lines (Comments): " & vbProject.CommentLines & _
        vbNewLine & vbNewLine & "           ------- References: -------"

    ReportPrjData2 = vbNewLine & "             ------- Imports: -------"

    With FrmResults.TreeView.Nodes
        For LoopVar = 1 To .Count
            If Left$(.Item(LoopVar).Key, 10) = "REFERENCE_" Then ReportPrjData = ReportPrjData & vbNewLine & .Item(LoopVar).Text
            If Left$(.Item(LoopVar).Key, 7) = "IMPORT_" Then ReportPrjData2 = ReportPrjData2 & vbNewLine & .Item(LoopVar).Text
        Next
    End With

    FrmReport.rtbReportText.Text = Replace$(FrmReport.rtbReportText.Text, "?PLACE>ProjectStats", ReportPrjData & vbNewLine & ReportPrjData2) ' Replace temp name in report with project stats

    FrmResults.TreeView.Nodes(1).Expanded = True
    Set vbProject = Nothing ' Delete created instance to prevent possible memory leaks
    IsScanning = False ' Set the scanning flag to false
    Exit Sub

NotVBNetProj:
    MsgBoxEx "不是一个VB.NET 工程文件!", vbCritical, "扫描错误", , , , , PicError, , 5
End Sub

Private Sub AddNETprojectToTreeview() ' Adds the .NET project's nodes to the treeview
    With FrmResults.TreeView.Nodes
        .Add 1, tvwChild, "NET_ProjectInfo_FileName", "Project FileName: " & ExtractFileName(FrmSelProject.cmbProjectPath.Text), "Info"
        .Add 1, tvwChild, "NET_ProjectInfo_Path", "Project Path: " & ExtractFilePath(FrmSelProject.cmbProjectPath.Text), "Info"
        .Add 1, tvwChild, "NET_ProjectInfo_Name", "Project Name: " & .Item(1).Text, "Info"

        .Add 1, tvwChild, "NET_AssembleInfo_Version", "Version: Unknown", "Info"

        .Add 1, tvwChild, "NETLines", "?", "Info"
        .Add 1, tvwChild, "NETLinesNB", "?", "Info"
        .Add 1, tvwChild, "NETLinesB", "?", "Info"
        .Add 1, tvwChild, "NETLinesComment", "?", "Info"

        .Add 1, tvwChild, "NETreferences", "References", "REFCOM"
        .Add 1, tvwChild, "NETimports", "Imports", "SysDLL"

        .Add 1, tvwChild, "NETfiles", "Files", "NETvb"
    End With
End Sub

Private Function GetNodeNum(NodeKey As String, Optional IndexNum As Long) As Long ' Find the node index of a item on the treeview,
    On Local Error Resume Next                                                                    ' from the given key.
    GetNodeNum = FrmResults.TreeView.Nodes.Item(NodeKey).Index
End Function

Private Sub LookAtDOTNETProjLine(linedata As String) ' Scans each project line
    Static LookingAtSettings As Boolean, SecondaryLineData As String
    linedata = Trim$(linedata)

    CurrentScanFile "NETProject" ' Change current scan item icon to a VB.NET file

    If linedata = "<Settings" Then LookingAtSettings = True: Exit Sub                                      ' \
    If LookingAtSettings = True And linedata = ">" Then LookingAtSettings = False: Exit Sub                ' | Retrieves settings from the
    If LookingAtSettings = True And Mid$(linedata, 1, 15) = "AssemblyName = " Then                          ' | .NET project file
        FrmResults.TreeView.Nodes.Add , , "NETproj", Mid$(linedata, 17, Len(linedata) - 17), "NETproject"   ' /
        AddNETprojectToTreeview
        Exit Sub
    End If

    If LookingAtSettings = True Then Exit Sub

    If Mid$(linedata, 1, 20) = "<Import Namespace = " Then AddImport Mid$(linedata, 22, Len(linedata) - 25): Exit Sub ' Get Imports from .NET project file
    If Mid$(linedata, 1, 10) = "<Reference" Then
        Line Input #ProjFileNum, linedata
        linedata = Trim$(linedata)
        AddReference Mid$(linedata, 9, Len(linedata) - 9) ' Retrieve references from the .NET project file
        Exit Sub
    End If

    If linedata = "<File" Then ' Get the filename of each .NET file and scan it
        Line Input #ProjFileNum, linedata
        Line Input #ProjFileNum, SecondaryLineData

        Do
            If InStrB(1, SecondaryLineData, "BuildAction = ") = 0 Then
                Line Input #ProjFileNum, SecondaryLineData
            Else
                Exit Do
            End If
        Loop

        If Trim$(SecondaryLineData) <> "BuildAction = ""Compile""" Then Exit Sub ' Make sure it's not a .TXT, .HTM, etc. file

        linedata = Trim$(linedata)
        linedata = Mid$(linedata, 12, Len(linedata) - 12)
        If InStrB(1, linedata, ".resx", vbTextCompare) = 0 Then AnalyseVBfile linedata      ' Scan the file
    End If
End Sub

Private Sub AnalyseVBfile(FileName As String) ' Look at a .NET file
    Dim linedata As String, LoopVar As Long, FileTitle As String

    CurrentScanFile "File" ' Change current scan item icon to a .NET file

    ScanningVBfile = False   '  \
    CodeDone = False         '  | Reset variables
    WaitForRegion = False    '  /

    ControlsInFile = 0

    If Not FileExists(FileName) Then ' Make sure the file exists
        If ShowFNFerrors Then MsgBoxEx "文件未找到: " & FileName, vbCritical, "扫描错误", , , , , PicError, , 5
        Exit Sub
    End If

    If InStrB(1, FileName, "AssemblyInfo.vb", vbTextCompare) Then ScanAssemblyFile FileName: Exit Sub  ' If the Assembly info file, look at it seperatly
    On Local Error Resume Next

    With FrmResults.TreeView.Nodes ' Add statistic nodes to treeview
        .Add GetNodeNum("NETfiles"), tvwChild, "TEMPvb", "?", "NETvb"
        .Add GetNodeNum("TEMPvb"), tvwChild, "TEMPvb_Lines", "?", "Info"
        .Add GetNodeNum("TEMPvb"), tvwChild, "TEMPvb_LinesNB", "?", "Info"
        .Add GetNodeNum("TEMPvb"), tvwChild, "TEMPvb_LinesB", "?", "Info"
        .Add GetNodeNum("TEMPvb"), tvwChild, "TEMPvb_LinesComment", "?", "Info"
        .Add GetNodeNum("TEMPvb"), tvwChild, "TEMPvb_Controls", "?", "Info"
        .Add GetNodeNum("TEMPvb"), tvwChild, "TEMPvb_Imports", "Imports", "SysDLL"
    End With

    Set vbFile = New ClsNETvbFile ' Create new instance

    VBFileNum = FreeFile
    Open FileName For Input As #VBFileNum

    Do
        Line Input #VBFileNum, linedata
        ScanVBfileLine Trim$(linedata) ' Look at each line
        If GetInputState Then DoEvents ' Process system events if there is messages in the keyboard/mouse buffers
    Loop While Not EOF(VBFileNum)

    Close #VBFileNum

    vbFile.Controls = ControlsInFile
    FileTitle = FrmResults.TreeView.Nodes(GetNodeNum("TEMPvb")).Text

    With FrmResults.TreeView.Nodes ' Fix the temp names and data
        .Item(GetNodeNum("TEMPvb_Imports")).Key = FileTitle & "_Imports"

        .Item(GetNodeNum("TEMPvb_Lines")).Text = "Lines (Inc. Blanks): " & (vbFile.CodeLines - 1)
        .Item(GetNodeNum("TEMPvb_Lines")).Key = FileTitle & "_Lines"
        .Item(GetNodeNum("TEMPvb_Controls")).Text = "Controls: " & vbFile.Controls
        .Item(GetNodeNum("TEMPvb_Controls")).Key = FileTitle & "_Controls"
        .Item(GetNodeNum("TEMPvb_LinesNB")).Text = "Lines (No Blanks): " & (vbFile.CodeLinesNB - 1)
        .Item(GetNodeNum("TEMPvb_LinesNB")).Key = FileTitle & "_LinesNB"
        .Item(GetNodeNum("TEMPvb_LinesB")).Text = "Lines (Blanks): " & vbFile.BlankLines
        .Item(GetNodeNum("TEMPvb_LinesB")).Key = FileTitle & "_LinesB"
        .Item(GetNodeNum("TEMPvb_LinesComment")).Text = "Lines (Comments): " & vbFile.CommentLines
        .Item(GetNodeNum("TEMPvb_LinesComment")).Key = FileTitle & "_LinesComment"
    End With

    For LoopVar = 1 To FrmResults.TreeView.Nodes.Count ' Fix all remaining temp key names
        If InStrB(1, FrmResults.TreeView.Nodes(LoopVar).Key, "TEMPvb") Then
            FrmResults.TreeView.Nodes(LoopVar).Key = Replace$(FrmResults.TreeView.Nodes(LoopVar).Key, "TEMPvb", FileTitle)
        End If
    Next

    AddReportText vbNewLine & "-------------------------------------------------------------"
    AddReportText "                   VISUAL BASIC .NET FILE"
    AddReportText "-------------------------------------------------------------"
    AddReportText "               File Name: " & ExtractFileName(FileName)
    AddReportText "               File Type: " & FileType
    AddReportText "                    Name: " & FileTitle
    AddReportText vbNewLine & "     Lines (Inc. Blanks): " & vbFile.CodeLines
    AddReportText "       Lines (No Blanks): " & vbFile.CodeLinesNB
    AddReportText "          Lines (Blanks): " & vbFile.BlankLines
    AddReportText "        Lines (Comments): " & vbFile.CommentLines

    If FileType = "Form" Then AddReportText vbNewLine & "                Controls: " & vbFile.Controls

    vbProject.BlankLines = vbFile.BlankLines     '  \
    vbProject.CodeLines = vbFile.CodeLines       '  | Add to the project's
    vbProject.CodeLinesNB = vbFile.CodeLinesNB   '  | total statistics
    vbProject.CommentLines = vbFile.CommentLines '  /

    FrmResults.TreeView.Nodes(GetNodeNum("NETfiles")).Expanded = True
    Set vbFile = Nothing ' Delete created instance to prevent possible memory leaks
End Sub

Private Sub AddImport(linedata As String) ' Add Imports to the treeview
    Dim NodeNum As Long

    NodeNum = GetNodeNum("NETimports")
    FrmResults.TreeView.Nodes.Add NodeNum, tvwChild, "IMPORT_" & linedata, linedata, "SysDLL"
End Sub

Private Sub AddReference(linedata As String) ' Add references to the treeview
    Dim NodeNum As Long

    If linedata = "c" Then Exit Sub

    NodeNum = GetNodeNum("NETreferences")
    FrmResults.TreeView.Nodes.Add NodeNum, tvwChild, "REFERENCE_" & linedata, linedata, "DLL"
End Sub


Private Sub ScanVBfileLine(linedata As String) ' Scan an individual line of a VB file
    vbProject.TotalLines = 1

RemoveIndent:                                                                                                           ' Loops here to remove any
    If Mid$(linedata, 1, 1) = vbTab Then linedata = Mid$(linedata, 5): GoTo RemoveIndent  ' indents (tabs) from the code

    If Mid$(linedata, 1, 9) = "End Class" Then CodeDone = True: Exit Sub

    If Mid$(linedata, 1, 48) = "#Region "" Windows Form Designer generated code """ Then WaitForRegion = True
    If Mid$(linedata, 1, 11) = "#End Region" Then WaitForRegion = False

    If Mid$(linedata, 1, 8) = "Imports " Then ' Add import data to treeview
        FrmResults.TreeView.Nodes.Add GetNodeNum("TEMPvb_Imports"), tvwChild, "TEMPvb_Import_" & Mid$(linedata, 9), Mid$(linedata, 9), "SysDLL"
        Exit Sub
    End If

    If Not ScanningVBfile Then GetNameAndType (linedata): Exit Sub ' Check what the file type is before scanning

    If CodeDone = True Then Exit Sub
    If WaitForRegion = True Then Exit Sub

    If IsHybridLine(linedata) = True Then ' Hybrid line (Both code and comment)
        vbFile.CommentLines = 1
        vbFile.CodeLines = 1
        vbFile.CodeLinesNB = 1
        Exit Sub
    End If

    If UCase$(Mid$(linedata, 1, 4)) = "REM " Then vbFile.CommentLines = 1: Exit Sub
    If Mid$(linedata, 1, 1) = "'" Then vbFile.CommentLines = 1: Exit Sub

    vbFile.CodeLines = 1
    If linedata = vbNullString Then vbFile.BlankLines = 1: Exit Sub
    vbFile.CodeLinesNB = 1
End Sub

Private Sub GetNameAndType(linedata As String) ' Retrieves the type of VB File (Form, etc.) and sets appropriate variables
    If Mid$(linedata, 1, 7) = "Module " Then ' >>MODULE<<
        FileType = "Module"

        FrmResults.TreeView.Nodes(GetNodeNum("TEMPvb")).Text = Mid$(linedata, 8)
        FrmResults.TreeView.Nodes(GetNodeNum("TEMPvb")).Image = "Module"
        ScanningVBfile = True
    End If

    If InStrB(1, linedata, ">") Then linedata = Trim$(Mid$(linedata, InStr(1, linedata, ">") + 1))

    If Left$(linedata, 13) = "Friend Class " Or Left$(linedata, 13) = "Public Class " Then
        FrmResults.TreeView.Nodes(GetNodeNum("TEMPvb")).Text = Mid$(linedata, 14)
        Line Input #VBFileNum, linedata

        If InStrB(1, linedata, "Inherits System.Windows.Forms.Form") Then  ' >>FORM<<
            FileType = "Form"

            FrmResults.TreeView.Nodes(GetNodeNum("TEMPvb")).Image = "Form"
        Else ' >>CLASS MODULE<<
            FileType = "Class Module"

            FrmResults.TreeView.Nodes(GetNodeNum("TEMPvb")).Image = "Class"
        End If
        
        ScanningVBfile = True
    End If
End Sub

Private Sub CurrentScanFile(FileType As String) ' Changes the little icon on the frmSelProject (if turned on) to
    If Not ShowCurrItemPic Then Exit Sub '              indicate what type of file is being scanned

    With FrmSelProject.imgCurrScanObjType
        Select Case FileType
            Case "NETProject"
                .Picture = FrmResults.ilstImages.ListImages("NETproject").Picture
            Case "File"
                .Picture = FrmResults.ilstImages.ListImages("NETvb").Picture
            Case "Clean"
                .Picture = FrmResults.ilstImages.ListImages(17).Picture
        End Select
        
        .Refresh
    End With
End Sub

Private Sub ScanAssemblyFile(FileName As String) ' Scans an Assembly.vb File to get project version
    Dim linedata As String, LinePos As Long, AssemblyFileNum As Integer

    AssemblyFileNum = FreeFile
    Open FileName For Input As #AssemblyFileNum

    Do
        Line Input #AssemblyFileNum, linedata
        If InStrB(1, linedata, "AssemblyVersion(") Then Exit Do
        DoEvents
    Loop While Not EOF(AssemblyFileNum)

    If linedata = vbNullString Then Exit Sub
    LinePos = InStr(1, linedata, "AssemblyVersion(") ' Find version line
    linedata = Mid$(linedata, LinePos + 17, Len(linedata) - LinePos - 19)
    If Right$(linedata, 1) = Chr$(34) Then linedata = Mid$(linedata, 1, Len(linedata) - 1)

    FrmResults.TreeView.Nodes.Item(GetNodeNum("NET_AssembleInfo_Version")).Text = "Version: " & linedata

    Close #AssemblyFileNum
End Sub
