Attribute VB_Name = "ModAnalyseVB6"
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:11:05
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：ModAnalyseVB6
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:11:05
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V4.12.0
'*************************************************************************

'-----------------------------------------------------------------------------------------------
'                                    VB6 SCANNER SYNOPSIS
'-----------------------------------------------------------------------------------------------
'   DeepLook is an advanced VB6 scanner. It is capable of showing almost all aspects of a source
'   project. The VB6 scanning engine in this module is (C) Dean Camera, and can show statistics
'   in a treeview and a text report. This module is capable of scanning projects made in all
'   versions of VB (except .NET) but is optimised and written around the VB6 version, and so
'   some errors may occur in projects saved in versions below VB5.
'
'   This VB scanner differs from many other scanners, because it excludes all header data from
'   the code statistics. Every VB file contains many hidden statements for the VB IDE, such as
'   control information on forms, and name info. Some other scanners include these lines when
'   they are not part of the actual code.
'
'   The VB scanning engine contained in this module is (C) Dean Camera.
'-----------------------------------------------------------------------------------------------

Option Explicit

'-----------------------------------------------------------------------------------------------
Private Project As ClsProjectFile
Private ProjectItem As ClsProjectItem

Private PrjItemPreviousLine As String

Private ShowSPFParams As Boolean
Private ShowFNFerrors As Integer
Private ShowIndividualLines As Boolean
Private CheckForMalicious As Long

Private RelatedDocumentFNames As String
Private TotalRelDocs As Long

Private InSub As Boolean, InFunction As Boolean, InProperty As Boolean
Private TotalSubs As Long, TotalFunctions As Long, TotalProperties As Long, TotalEvents As Long, TotalDecSubs As Long, TotalDecFunctions As Long
Private CurrSPFLines As Long, CurrSPFLinesNB As Long, CurrSPFName As String, CurrSPFColour As Long
Private GroupTotalLines As Long, GroupTotalLinesNB As Long, GroupTotalPrj As Long

Private FileInfo As ClsFileProp
Private ThisPrgNodeStart As Long
Private NodeNum As Long

Private Group1Percent As Single, Project1Percent As Single
Private VBFileNum As Integer

Private MaliciousKeywordsBuffer() As String
Private MaliciousKeywordsBuffer_Elements As Long

Public UsedFonts() As String ' Must be public, accessed by ModFileSeachHandler
Private PMaliciousCode() As String
'-----------------------------------------------------------------------------------------------

Public Sub ClearArrays() 'Reset the variables and clear arrays ready for another project to be scanned
    ModVariableHandler.ClearGlobals

    UsesGroup = False
    InSub = False
    InFunction = False
    InProperty = False
    RelatedDocumentFNames = ""
    FilesRootDirectory = ""

    Set FileInfo = New ClsFileProp

    ReDim PMaliciousCode(0) As String ' Potentially Malicious Code is stored in an array for later retrieval and sorting (this clears it)

    ShowSPFParams = GetSetting("DeepLook", "Options", "ShowSPFParams", 1)      ' \
    CheckForMalicious = GetSetting("DeepLook", "Options", "PMCCheck", 1)       ' |
    ShowCurrItemPic = GetSetting("DeepLook", "Options", "ShowCurrItemPic", 1)  ' | Get DeepLook's settings from the registry
    ShowFNFerrors = GetSetting("DeepLook", "Options", "ShowFNFErrors", 1)      ' |
    ShowIndividualLines = GetSetting("DeepLook", "Options", "ShowSPFLines", 0) ' /

    If ShowCurrItemPic Then                              '  \
        FrmSelProject.pgbAPB.Width = 4215                '  |
        FrmSelProject.imgCurrScanObjType.Visible = True  '  |  Rearrange items on the SelProject form
    Else                                                 '  |  depending on the program's settings
        FrmSelProject.pgbAPB.Width = 4695                '  |
        FrmSelProject.imgCurrScanObjType.Visible = False '  |
    End If                                               '  /
End Sub

Public Sub PostScanDeleteObjects() ' Delete all the created objects after scanning - prevents possible leaks
    Set FileInfo = Nothing
    Set Project = Nothing
    Set ProjectItem = Nothing
    Set FileInfo = Nothing
End Sub

'-----------------------------------------------------------------------------------------------

Public Sub AnalyseGroup(GroupFileName As String) ' Master command to analyse a Group (.vbg) file.
    Dim linedata As String, DoneGrp As Long, BlankLines As Long
    Dim GroupFileNum As Integer

    IsScanning = True ' Set the scanning flag to true

    CurrentScanFile "Group" ' Change small picture of current operation to a VB Group symbol
    DoEvents

    FilesRootDirectory = GetRootDirectory(GroupFileName) ' Retrieve directory of the group file

    FrmSelProject.lblScanningName.Caption = ExtractFileName(GroupFileName) 'Change the label showing the current file name to the group's file name

    GroupFileName = FixRelPath(GroupFileName)

    If Not FileExists(GroupFileName) Then ' Throw a graceful error if the file is not found
        MsgBoxEx "工程组文件未找到!", vbCritical, "扫描错误", , , , , PicError, , 5
        IsScanning = False ' Set the scanning flag to false
        Exit Sub
    End If

    With FrmResults.TreeView.Nodes
        .Add , , "GROUP", "工程组", "Group" 'Add the group to the treeview
        .Add 1, tvwChild, "GROUP_Lines", "?", "Info"
        .Add 1, tvwChild, "GROUP_LinesNB", "?", "Info"
        .Add 1, tvwChild, "GROUP_TotalPrj", "?", "Info"
    End With

    GroupTotalPrj = 0
    UsesGroup = True ' Tell the AnalyseProject sub that it should be adding project files to the group node

    GroupTotalLines = 0
    GroupTotalLinesNB = 0

    GroupFileNum = FreeFile
    Open GroupFileName For Input As #GroupFileNum

    AddReportText vbNewLine & "*************************************************************"
    AddReportText "**********************Visual Basic Group*********************"
    AddReportText "*************************************************************"
    AddReportText "                   File Name: " & ExtractFileName(GroupFileName)
    AddReportText "*************************************************************"
    AddReportText "       ?PLACE>GRPTotalProjects"
    AddReportText " ?PLACE>GRPTotalLines"
    AddReportText "   ?PLACE>GRPTotalLinesNB"
    AddReportText "*************************************************************"

    Line Input #GroupFileNum, linedata ' Get the line data and increment the completed percentage
    DoneGrp = Len(linedata) + 2

    If InStrB(1, linedata, "VBGROUP") = 0 Then GoTo InvalidGroupFile ' Group file must contain valid VB6 Group Header

    Group1Percent = (100 / LOF(GroupFileNum))

    Do While Not EOF(GroupFileNum)
        FrmSelProject.pgbAPB.Value2 = Group1Percent * DoneGrp     ' Increment progress bar
        Line Input #GroupFileNum, linedata ' Retrieve line data
        DoneGrp = DoneGrp + Len(linedata)

        If GetInputState Then DoEvents ' Process system events if there is messages in the keyboard/mouse buffers

        If IsExit Then Exit Do ' If quitting, don't process any more of the file

        If Left$(linedata, 15) = "StartupProject=" Then
            AnalyseVBProject Mid$(linedata, 16), True

            If Not Project Is Nothing Then
                GroupTotalLines = GroupTotalLines + Project.ProjectLines
                GroupTotalLinesNB = GroupTotalLinesNB + Project.ProjectLinesNB  'Pass a TRUE value to the "StartupProjectInGroup" parameter
            End If
        ElseIf Left$(linedata, 8) = "Project=" Then
            AnalyseVBProject Mid$(linedata, 9)

            If Not Project Is Nothing Then
                GroupTotalLines = GroupTotalLines + Project.ProjectLines
                GroupTotalLinesNB = GroupTotalLinesNB + Project.ProjectLinesNB
            End If
        End If
    Loop

    Close #GroupFileNum

    Set Project = Nothing ' Kill the project statistics object once all processing of the project is completed
    FrmResults.btnFileCopy.Enabled = False

    With FrmResults.TreeView.Nodes
        .Item(2).Text = "总计代码行 (包括空行): " & Format$(GroupTotalLines, "###,###,###")
        .Item(3).Text = "总计代码行(不包括空行): " & Format$(GroupTotalLinesNB, "###,###,###")
        .Item(4).Text = "总计工程组: " & Format$(GroupTotalPrj, "###,###,###")

        FrmReport.rtbReportText.Text = Replace$(FrmReport.rtbReportText.Text, "?PLACE>GRPTotalProjects", "Total Project Files: " & Format$(GroupTotalPrj, "###,###,###"), 1, 1)
        FrmReport.rtbReportText.Text = Replace$(FrmReport.rtbReportText.Text, "?PLACE>GRPTotalLines", "Total Lines (Inc. Blanks): " & Format$(GroupTotalLines, "###,###,###"), 1, 1)
        FrmReport.rtbReportText.Text = Replace$(FrmReport.rtbReportText.Text, "?PLACE>GRPTotalLinesNB", "Total Lines (No Blanks): " & Format$(GroupTotalLinesNB, "###,###,###"), 1, 1)

        .Item(1).Expanded = True ' Expand the group node on the treeview

        BlankLines = (GroupTotalLines - GroupTotalLinesNB)

        With FrmResults.chtGroupChart
            .Column = 1
            .Data = Round((100 / OneIfNull(GroupTotalLines)) * OneIfNull(GroupTotalLines - BlankLines))
            .Column = 2
            .Data = Round((100 / OneIfNull(GroupTotalLines)) * OneIfNull(BlankLines))
        End With
    End With

    IsScanning = False ' No longer scanning a file
    Exit Sub

InvalidGroupFile:
    Close #GroupFileNum
    MsgBoxEx "无效的工程组文件 - 文件不是一个 VB6 工程组!", vbCritical, "扫描错误", , , , , PicError, "Oops!|", 5
End Sub

Public Sub AnalyseVBProject(ProjectFileName As String, Optional StartupProjectInGroup As Boolean) 'Master command to analyse a VB Project
    Dim linedata As String, DonePrj As Long, BlankLines As Long
    Dim ProjFileNum As Integer

    IsScanning = True ' Set the scanning flag to true

    With FrmSelProject
        .pgbAPB.Value = 0
        .lblScanPhase.Caption = "扫描代码阶段"
        .lblScanPhase.ForeColor = &HA000&
        .pgbAPB.Color = &HC000&
    End With

    ModVariableHandler.ClearGlobals ' Clear found Global Variables
    ModVariableHandler.ClearLocals ' Clear found Local Variables

    PROJECTMODE = VB6
    GroupTotalPrj = GroupTotalPrj + 1 ' Increment total projects scanned (for group scan)

    CurrentScanFile "Project" ' Change small picture of current operation to a VB Project symbol

    FrmSelProject.lblScanningName.Caption = ExtractFileName(ProjectFileName) 'Change the label on the SelProject form to the project's file name

    DoEvents ' Prevent VB from locking up when scanning projects - and so the progressbar can refresh

    ProjectFileName = FixRelPath(ProjectFileName)
    ProjFileNum = FreeFile

    If Not FileExists(FilesRootDirectory & ProjectFileName) And Not FileExists(ProjectFileName) Then
        If UsesGroup Then
            If ShowFNFerrors And Not IsExit Then MsgBoxEx "工程文件 """ & ProjectFileName & """ 没有找到!", vbCritical, "扫描错误", , , , , PicError, "Oops!|", 5 'Throw an error if the file cannot be found (only shown if show FNF errors setting turned on)

            FrmResults.TreeView.Nodes.Add 1, tvwChild, "PROJECT_" & ProjectFileName, ProjectFileName, "Unknown"
        Else
            MsgBoxEx "工程文件 """ & ProjectFileName & """ 未发现!", vbCritical, "扫描错误", , , , , PicError, "Oops!|", 5 'Throw an error if the file cannot be found (always shown if project not part of a group)

            FrmResults.TreeView.Nodes.Add , , "PROJECT_?", "?", "Project" 'File found, so add to treeview
        End If

        IsScanning = False ' Set the scanning flag to false

        Exit Sub
    End If

    Set Project = New ClsProjectFile ' Initialise a new instance of ClsProjectFile to store project statistics/data

    ThisPrgNodeStart = FrmResults.TreeView.Nodes.Count ' Recode the node start of the project to increase cleanup time

    Project.ProjectPath = ProjectFileName

    AddProjectToTreeView
    If StartupProjectInGroup Then FrmResults.TreeView.Nodes(GetNodeNum("PROJECT_?")).Bold = True 'Make the project bold in the treeview if it is a startup project

    AddReportText vbNewLine & "============================================================="
    AddReportText "===================Visual Basic Project File================="
    AddReportText "============================================================="
    AddReportText "               File Name: " & ExtractFileName(ProjectFileName)
    AddReportText "            Project Name: " & "?PLACE>ProjectName"
    AddReportText "         Project Version: " & "?PLACE>ProjectVersion"
    AddReportText "============================================================="
    AddReportText "?PLACE>ProjectStats"

    FrmResults.btnFileCopy.Enabled = True

    If Not FileExists(FilesRootDirectory & ProjectFileName) Then
        Open ProjectFileName For Input As #ProjFileNum
    Else
        Open FilesRootDirectory & ProjectFileName For Input As #ProjFileNum
    End If

    Project1Percent = (100 / LOF(ProjFileNum))

    Do While Not EOF(ProjFileNum)
        If IsExit Then Exit Do ' If quitting, don't process any more of the file

        FrmSelProject.pgbAPB.Value = Project1Percent * DonePrj ' Increase progressbar
        Line Input #ProjFileNum, linedata ' Get the line data from the project file
        AnalyseProjectLine linedata ' Look at the project line to see what it contains
        DonePrj = DonePrj + Len(linedata)
        DoEvents ' Process system events
    Loop

    Close #ProjFileNum

    AddFontsToTreeview ' Add used fonts to the treeview

    GetProjectEXEStats ' Get statistics on the compiled project EXE (if it exists)

    If EXENewOrOld = "1N" Then ' Project file newer than EXE
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?"), tvwChild, "PROJECT_?_EXEFileVerDetails", "The Project File is a newer version than your compiled EXE. Please Recompile.", "Warning"
    ElseIf EXENewOrOld = "2N" Then ' EXE newer than Project file
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?"), tvwChild, "PROJECT_?_EXEFileVerDetails", "The compiled EXE is a newer version than your Project File.", "Warning"
    End If

    FixPrjTreeViewItems ' Fix up the treeview items, captions and keys. All items use a "?" in the key instead of the project name
    '                     until the FixPrjTreeViewItems sub changes this to prevent group projects from interfering with each other.

    AddProjectReportText ' Add collected stats about the current project to the report

    ModVariableHandler.AnalyseVBProjectForVars ProjectFileName ' Scan the project for unused variables

    If GetNodeNum("GROUP") = 0 Then ' Single project scan (not part of a group)
        FrmResults.TreeView.Nodes(1).Expanded = True ' If not part of a group file, expand the project

        BlankLines = (Project.ProjectLines - Project.ProjectLinesNB)

        With FrmResults.chtGroupChart
            .Column = 1
            .Data = Round((100 / OneIfNull(Project.ProjectLines)) * OneIfNull(Project.ProjectLines - BlankLines))
            .Column = 2
            .Data = Round((100 / OneIfNull(Project.ProjectLines)) * OneIfNull(BlankLines))
        End With

        With FrmResults.chtProjectChart
            .Column = 1
            .Data = Round((100 / OneIfNull(Project.ProjectLines + Project.ProjectCommentLines)) * OneIfNull(Project.ProjectLines - BlankLines))
            .Column = 2
            .Data = Round((100 / OneIfNull(Project.ProjectLines + Project.ProjectCommentLines)) * OneIfNull(BlankLines))
            .Column = 3
            .Data = Round((100 / OneIfNull(Project.ProjectLines + Project.ProjectCommentLines)) * OneIfNull(Project.ProjectCommentLines))
        End With

        FrmResults.lblSelProjName.Caption = Project.ProjectName

        Set Project = Nothing ' Kill the project statistics object once all processing of the project is completed
        IsScanning = False ' No longer scanning a file
    End If
End Sub

Public Sub AnalyseSingleVBItem(FileName As String)  ' Master sub to analyse a single file (.fmr, .bas, etc.) rather than a project
    Dim ItemKeyNum As Long, QuestionPos As Long

    IsScanning = True ' Set the scanning flag to true
    PROJECTMODE = VB6
    FileName = FixRelPath(FileName)

    With FrmSelProject
        .pgbAPB.Value = 0
        .lblScanPhase.Caption = "扫描代码阶段"
        .lblScanPhase.ForeColor = &HA000&
        .pgbAPB.Color = &HC000&
    End With

    DoEvents

    FrmResults.btnFileCopy.Enabled = False

    If Not FileExists(FileName) Then ' If file doesn't exist, throw an error
        MsgBoxEx "文件未发现!", vbCritical, "扫描错误", , , , , PicError
        IsScanning = False ' Set the scanning flag to false
        Exit Sub
    End If

    Set Project = New ClsProjectFile ' Create a new project class to stop the analysis subs from giving errors
    '                                  when trying to add to it's values - this is not actually used in the treeview

    Set ProjectItem = New ClsProjectItem ' Create a new file instance to hold file statistics

    With FrmResults.TreeView.Nodes

        FrmResults.TreeView.Nodes.Add , , "PROJECT_?", "(Temp Project)", "Project"

        Select Case Right$(UCase$(FileName), 4) ' Add the required nodes depending on the file type
            Case ".FRM"
                .Add GetNodeNum("PROJECT_?"), tvwChild, "PROJECT_?_FORMS", "Forms", "Form"
                .Add GetNodeNum("PROJECT_?_FORMS"), tvwChild, "PROJECT_?_FORMS_LINES", "0"
                .Add GetNodeNum("PROJECT_?_FORMS"), tvwChild, "PROJECT_?_FORMS_VARIABLES", "0"
                GetFormStats FileName
                .Remove GetNodeNum("PROJECT_?_FORMS_LINES")
                .Remove GetNodeNum("PROJECT_?_FORMS_VARIABLES")
            Case ".BAS"
                .Add GetNodeNum("PROJECT_?"), tvwChild, "PROJECT_?_MODULES", "Modules", "Module"
                .Add GetNodeNum("PROJECT_?_MODULES"), tvwChild, "PROJECT_?_MODULES_LINES", "0"
                .Add GetNodeNum("PROJECT_?_MODULES"), tvwChild, "PROJECT_?_MODULES_VARIABLES", "0"
                GetModuleStats FileName & ";" & ExtractFileName(FileName)
                .Remove GetNodeNum("PROJECT_?_MODULES_LINES")
                .Remove GetNodeNum("PROJECT_?_MODULES_VARIABLES")
            Case ".CLS"
                .Add GetNodeNum("PROJECT_?"), tvwChild, "PROJECT_?_CLASSES", "Classes", "Class"
                .Add GetNodeNum("PROJECT_?_CLASSES"), tvwChild, "PROJECT_?_CLASSES_LINES", "0"
                .Add GetNodeNum("PROJECT_?_CLASSES"), tvwChild, "PROJECT_?_CLASSES_VARIABLES", "0"
                GetClassStats FileName & ";" & ExtractFileName(FileName)
                .Remove GetNodeNum("PROJECT_?_CLASSES_LINES")
                .Remove GetNodeNum("PROJECT_?_CLASSES_VARIABLES")
            Case ".CTL"
                .Add GetNodeNum("PROJECT_?"), tvwChild, "PROJECT_?_USERCONTROLS", "User Controls", "UserControl"
                .Add GetNodeNum("PROJECT_?_USERCONTROLS"), tvwChild, "PROJECT_?_USERCONTROLS_LINES", "0"
                .Add GetNodeNum("PROJECT_?_USERCONTROLS"), tvwChild, "PROJECT_?_USERCONTROLS_VARIABLES", "0"
                GetUserControlStats FileName
                .Remove GetNodeNum("PROJECT_?_USERCONTROLS_LINES")
                .Remove GetNodeNum("PROJECT_?_USERCONTROLS_VARIABLES")
            Case ".PAG"
                .Add GetNodeNum("PROJECT_?"), tvwChild, "PROJECT_?_PROPERTYPAGES", "Property Pages", "PropertyPage"
                .Add GetNodeNum("PROJECT_?_PROPERTYPAGES"), tvwChild, "PROJECT_?_PROPERTYPAGES_LINES", "0"
                .Add GetNodeNum("PROJECT_?_PROPERTYPAGES"), tvwChild, "PROJECT_?_PROPERTYPAGES_VARIABLES", "0"
                GetPropertyPageStats FileName
                .Remove GetNodeNum("PROJECT_?_PROPERTYPAGES_LINES")
                .Remove GetNodeNum("PROJECT_?_PROPERTYPAGES_VARIABLES")
            Case ".DOB"
                .Add GetNodeNum("PROJECT_?"), tvwChild, "PROJECT_?_USERDOCUMENTS", "User Documents", "UserDocument"
                .Add GetNodeNum("PROJECT_?_USERDOCUMENTS"), tvwChild, "PROJECT_?_USERDOCUMENTS_LINES", "0"
                .Add GetNodeNum("PROJECT_?_USERDOCUMENTS"), tvwChild, "PROJECT_?_USERDOCUMENTS_VARIABLES", "0"
                GetUserDocumentStats FileName
                .Remove GetNodeNum("PROJECT_?_USERDOCUMENTS_LINES")
                .Remove GetNodeNum("PROJECT_?_USERDOCUMENTS_VARIABLES")
            Case ".DSR"
                .Add GetNodeNum("PROJECT_?"), tvwChild, "PROJECT_?_DESIGNERS", "Designers", "Designer"
                .Add GetNodeNum("PROJECT_?_DESIGNERS"), tvwChild, "PROJECT_?_DESIGNERS_LINES", "0"
                .Add GetNodeNum("PROJECT_?_DESIGNERS"), tvwChild, "PROJECT_?_DESIGNERS_VARIABLES", "0"
                GetDesignerStats FileName
                .Remove GetNodeNum("PROJECT_?_DESIGNERS_LINES")
                .Remove GetNodeNum("PROJECT_?_DESIGNERS_VARIABLES")
        End Select

        With FrmSelProject
            .pgbAPB.Value = 100
            DoEvents

            .pgbAPB.Value = 0
            .pgbAPB.Color = 8421631
            .lblScanPhase.Caption = "未用变量分析阶段"
            .lblScanPhase.ForeColor = 8421631
            .imgCurrScanObjType.Picture = FrmResults.ilstImages.ListImages(29).Picture
            DoEvents

            ModVariableHandler.ClearLocals ' Initialise the unused variable scanner
            ModVariableHandler.AnalyseVBProjectForVars FileName ' Scan the file for unused variables

            .pgbAPB.Value = 100
            DoEvents
        End With

        For ItemKeyNum = 1 To .Count ' Go over every node, checking for Related Documents, and replacing the temp project name "?" with the correct
            .Item(ItemKeyNum).Expanded = True ' Expand the node

            Do
                QuestionPos = InStr(1, .Item(ItemKeyNum).Text, "  ") ' Checks for indents, which make line continuations look wrong
                If QuestionPos Then  ' "QuestionPos" variable used here only to save on variable requirements
                    .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, QuestionPos - 1) & Mid$(.Item(ItemKeyNum).Text, QuestionPos + 2)
                Else ' No indents left - exit the loop
                    Exit Do
                End If
            Loop

            If InStrB(1, .Item(ItemKeyNum).Text, "[EXT]") Then  'Is an external (DLL) call (declared sub or function)
                .Item(ItemKeyNum).Parent.ForeColor = RGB(150, 150, 150) 'Make key text a light-grey colour
                .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 10) 'Remove the now unnecessary "[EXT]" text
            ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[1]" Then 'SPF Colour 1
                .Item(ItemKeyNum).Parent.ForeColor = RGB(130, 0, 200) 'Colour Key Text
                .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
            ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[2]" Then 'SPF Colour 2
                .Item(ItemKeyNum).Parent.ForeColor = RGB(200, 0, 150) 'Colour Key Text
                .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
            ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[3]" Then 'SPF Colour 3
                .Item(ItemKeyNum).Parent.ForeColor = RGB(10, 150, 10) 'Colour Key Text
                .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
            ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[6]" Then 'SPF Colour 4
                .Item(ItemKeyNum).Parent.ForeColor = RGB(249, 164, 0) 'Colour Key Text
                .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
            ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[7]" Then 'SPF Colour 5
                .Item(ItemKeyNum).Parent.ForeColor = RGB(217, 206, 19) 'Colour Key Text
                .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
            ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[8]" Then 'SPF Colour 6
                .Item(ItemKeyNum).Parent.ForeColor = RGB(19, 217, 192) 'Colour Key Text
                .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
            ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[4]" Then 'SPF Colour 7
                .Item(ItemKeyNum).Parent.ForeColor = RGB(20, 90, 100) 'Colour Key Text
                .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
            ElseIf Right$(.Item(ItemKeyNum).Text, 3) = "[5]" Then 'SPF Colour 8
                .Item(ItemKeyNum).Parent.ForeColor = RGB(50, 23, 80) 'Colour Key Text
                .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
            End If

            If InStrB(1, .Item(ItemKeyNum).Text, "No Blanks):") Then
                If Right$(.Item(ItemKeyNum).Text, 2) = "0" Then
                    If InStrB(1, .Item(ItemKeyNum).Parent.Key, "_SUB") Then
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(255, 0, 0)
                    ElseIf InStrB(1, .Item(ItemKeyNum).Parent.Key, "_FUNC") Then
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(255, 0, 0)
                    ElseIf InStrB(1, .Item(ItemKeyNum).Parent.Key, "_PROP") Then
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(255, 0, 0)
                    End If
                End If
            End If
        Next
    End With

    With FrmResults.chtGroupChart
        .Column = 1
        .Data = 0
        .Column = 2
        .Data = 0
    End With

    Set Project = Nothing ' Kill the project statistics object once all processing of the project is completed
    Set ProjectItem = Nothing ' Kill the project item statistics object once all processing of the file is completed

    IsScanning = False ' Set the scanning flag to false
End Sub

Private Sub AnalyseProjectLine(linedata As String)  ' Sub to look at the contents of the project's line and analyse accordingly
    Dim VariableName As String, VariableValue As String

    CurrentScanFile "Project" ' Change the small picture to the project symbol

    ' The following code identifies, cuts and parses project data lines to the appropriate subroutine for scanning.
    ' The order of the IFs are important; the most commonly used lines are checked first to reduce scanning time.

    If InStrB(1, linedata, "=") = 0 Then Exit Sub

    VariableName = Left$(linedata, InStr(1, linedata, "="))
    VariableValue = Mid$(linedata, InStr(1, linedata, "=") + 1)

    If VariableName = "Form=" Then
        Set ProjectItem = New ClsProjectItem
        GetFormStats VariableValue
        Set ProjectItem = Nothing
    ElseIf VariableName = "Module=" Then
        Set ProjectItem = New ClsProjectItem
        GetModuleStats VariableValue
        Set ProjectItem = Nothing
    ElseIf VariableName = "Class=" Then
        Set ProjectItem = New ClsProjectItem
        GetClassStats VariableValue
        Set ProjectItem = Nothing
    ElseIf VariableName = "UserControl=" Then
        Set ProjectItem = New ClsProjectItem
        GetUserControlStats VariableValue
        Set ProjectItem = Nothing
    ElseIf VariableName = "PropertyPage=" Then
        Set ProjectItem = New ClsProjectItem
        GetPropertyPageStats VariableValue
        Set ProjectItem = Nothing
    ElseIf VariableName = "UserDocument=" Then
        Set ProjectItem = New ClsProjectItem
        GetUserDocumentStats VariableValue
        Set ProjectItem = Nothing
    ElseIf VariableName = "Designer=" Then
        Set ProjectItem = New ClsProjectItem
        GetDesignerStats VariableValue
        Set ProjectItem = Nothing

    ElseIf VariableName = "ResFile32=" Then
        GetRelatedDocStats Mid$(VariableValue, 2, Len(VariableValue) - 2)
    ElseIf VariableName = "RelatedDoc=" Then
        GetRelatedDocStats VariableValue
    ElseIf VariableName = "Object=" Then
        GetComponentStats VariableValue
    ElseIf VariableName = "Reference=" Then
        GetReferenceStats VariableValue

    ElseIf VariableName = "Name=" Then
        Project.ProjectName = Mid$(VariableValue, 2, Len(VariableValue) - 2)
    ElseIf VariableName = "Title=" Then
        Project.ProjectTitle = Mid$(VariableValue, 2, Len(VariableValue) - 2)
    ElseIf VariableName = "Type=" Then
        Project.ProjectProjectType = VariableValue
    ElseIf VariableName = "Startup=" Then
        Project.ProjectStartupItem = Mid$(VariableValue, 2, Len(VariableValue) - 2)

    ElseIf VariableName = "MajorVer=" Then
        Project.ProjectVersion = VariableValue & "."
    ElseIf VariableName = "MinorVer=" Then
        Project.ProjectVersion = VariableValue & "."
    ElseIf VariableName = "RevisionVer=" Then
        Project.ProjectVersion = VariableValue

    ElseIf VariableName = "ExeName32=" Then
        Project.ProjectEXEFName = Mid$(VariableValue, 2, Len(VariableValue) - 2)
    ElseIf VariableName = "Path32=" Then
        Project.ProjectEXEPath = Mid$(VariableValue, 2, Len(VariableValue) - 2)
    End If
End Sub

Private Function GetNodeNum(NodeKey As String) As Long  'Find the node index of a item on the treeview,
    On Local Error Resume Next ' Prevents errors when the node cannot be found                                                                    'from the given key.
    GetNodeNum = FrmResults.TreeView.Nodes.Item(NodeKey).Index
End Function

Private Sub AddProjectToTreeView() ' Sub to add all the standard nodes to the treeview (Forms, Project Info, etc.)
    Dim ProjectRootNode As Long, SPFNode As Long

    If UsesGroup Then ' If the project is part of a group, add it to the group node that was added by the Group scanning sub
        FrmResults.TreeView.Nodes.Add 1, tvwChild, "PROJECT_?", "?", "Project"
    Else ' Add it as the first node if not part of a group
        FrmResults.TreeView.Nodes.Add , , "PROJECT_?", "?", "Project"
    End If

    ProjectRootNode = GetNodeNum("PROJECT_?") ' Save time by finding the project's node index only once

    With FrmResults.TreeView.Nodes ' Looks better with a "With" statement
        .Add ProjectRootNode, tvwChild, "PROJECT_?_TITLE", "?", "Info"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_VERSION", "?", "Info"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_TYPE", "?", "Info"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_STARTUPITEM", "?", "Info"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_SOURCESAFE", "源代码保护: " & CheckIsSourceSafe(Project.ProjectPath), "Info"

        .Add ProjectRootNode, tvwChild, "PROJECT_?_TOTALS", "工程统计", "ProjStats"
        .Add GetNodeNum("PROJECT_?_TOTALS"), tvwChild, "PROJECT_?_LINES", "?", "Info"
        .Add GetNodeNum("PROJECT_?_TOTALS"), tvwChild, "PROJECT_?_LINESNB", "?", "Info"
        .Add GetNodeNum("PROJECT_?_TOTALS"), tvwChild, "PROJECT_?_LINESCOMMENT", "?", "Info"
        .Add GetNodeNum("PROJECT_?_TOTALS"), tvwChild, "PROJECT_?_VARIABLES", "?", "Info"
        .Add GetNodeNum("PROJECT_?_TOTALS"), tvwChild, "PROJECT_?_CONSTANTS", "?", "Info"
        .Add GetNodeNum("PROJECT_?_TOTALS"), tvwChild, "PROJECT_?_TYPES", "?", "Info"
        .Add GetNodeNum("PROJECT_?_TOTALS"), tvwChild, "PROJECT_?_ENUMS", "?", "Info"

        .Add ProjectRootNode, tvwChild, "PROJECT_?_PROJINFO", "工程文件信息", "ProjFileInfo"

        FileInfo.FindFileInfo Project.ProjectPath, False
        .Add GetNodeNum("PROJECT_?_PROJINFO"), tvwChild, "PROJECT_?_PROJINFO_MODIFIED", "最近修改日期: " & FileInfo.LastWriteTime, "Info"
        .Add GetNodeNum("PROJECT_?_PROJINFO"), tvwChild, "PROJECT_?_PROJINFO_ACCESS", "最近访问日期: " & FileInfo.LastAccessTime, "Info"

        .Add ProjectRootNode, tvwChild, "PROJECT_?_FORMS", "窗体", "Form"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_MODULES", "模块", "Module"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_CLASSES", "类模块", "Class"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_USERCONTROLS", "用户控件", "UserControl"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_USERDOCUMENTS", "用户文档", "UserDocument"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_PROPERTYPAGES", "属性页", "PropertyPage"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_DESIGNERS", "设计页", "Designer"

        .Add GetNodeNum("PROJECT_?_FORMS"), tvwChild, "PROJECT_?_FORMS_COUNT", "?", "Info"
        .Add GetNodeNum("PROJECT_?_MODULES"), tvwChild, "PROJECT_?_MODULES_COUNT", "?", "Info"
        .Add GetNodeNum("PROJECT_?_CLASSES"), tvwChild, "PROJECT_?_CLASSES_COUNT", "?", "Info"
        .Add GetNodeNum("PROJECT_?_USERCONTROLS"), tvwChild, "PROJECT_?_USERCONTROLS_COUNT", "?", "Info"
        .Add GetNodeNum("PROJECT_?_USERDOCUMENTS"), tvwChild, "PROJECT_?_USERDOCUMENTS_COUNT", "?", "Info"
        .Add GetNodeNum("PROJECT_?_PROPERTYPAGES"), tvwChild, "PROJECT_?_PROPERTYPAGES_COUNT", "?", "Info"
        .Add GetNodeNum("PROJECT_?_DESIGNERS"), tvwChild, "PROJECT_?_DESIGNERS_COUNT", "?", "Info"

        .Add GetNodeNum("PROJECT_?_FORMS"), tvwChild, "PROJECT_?_FORMS_LINES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_MODULES"), tvwChild, "PROJECT_?_MODULES_LINES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_CLASSES"), tvwChild, "PROJECT_?_CLASSES_LINES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_USERCONTROLS"), tvwChild, "PROJECT_?_USERCONTROLS_LINES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_USERDOCUMENTS"), tvwChild, "PROJECT_?_USERDOCUMENTS_LINES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_PROPERTYPAGES"), tvwChild, "PROJECT_?_PROPERTYPAGES_LINES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_DESIGNERS"), tvwChild, "PROJECT_?_DESIGNERS_LINES", "0", "Info"

        .Add GetNodeNum("PROJECT_?_FORMS"), tvwChild, "PROJECT_?_FORMS_VARIABLES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_MODULES"), tvwChild, "PROJECT_?_MODULES_VARIABLES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_CLASSES"), tvwChild, "PROJECT_?_CLASSES_VARIABLES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_USERCONTROLS"), tvwChild, "PROJECT_?_USERCONTROLS_VARIABLES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_USERDOCUMENTS"), tvwChild, "PROJECT_?_USERDOCUMENTS_VARIABLES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_PROPERTYPAGES"), tvwChild, "PROJECT_?_PROPERTYPAGES_VARIABLES", "0", "Info"
        .Add GetNodeNum("PROJECT_?_DESIGNERS"), tvwChild, "PROJECT_?_DESIGNERS_VARIABLES", "0", "Info"

        .Add ProjectRootNode, tvwChild, "PROJECT_?_RELATEDDOCUMENTS", "相关文档", "RelatedDocuments"
        .Add GetNodeNum("PROJECT_?_RELATEDDOCUMENTS"), tvwChild, "PROJECT_?_RELATEDDOCUMENTS_COUNT", "总计:", "Info"
        .Add ProjectRootNode, tvwChild, "PROJECT_?_REFCOM", "引用 & 组件", "REFCOM"
        .Add GetNodeNum("PROJECT_?_REFCOM"), tvwChild, "PROJECT_?_REFCOM_COUNT", "总计:", "Info"

        .Add ProjectRootNode, tvwChild, "PROJECT_?_DECDLLS", "声明 DLLs", "DLL"

        .Add ProjectRootNode, tvwChild, "PROJECT_?_SPF", "过程, 函数和属性", "SPF"
        SPFNode = GetNodeNum("PROJECT_?_SPF")

        .Add SPFNode, tvwChild, "PROJECT_?_SPF_SUBS", "0", "Info"
        .Add SPFNode, tvwChild, "PROJECT_?_SPF_FUNCTIONS", "0", "Info"
        .Add SPFNode, tvwChild, "PROJECT_?_SPF_PROPERTIES", "0", "Info"
        .Add SPFNode, tvwChild, "PROJECT_?_SPF_EVENTS", "0", "Info"
        .Add SPFNode, tvwChild, "PROJECT_?_SPF_DECLAREDSUBS", "0", "Info"
        .Add SPFNode, tvwChild, "PROJECT_?_SPF_DECLAREDFUNCTIONS", "0", "Info"
        .Add SPFNode, tvwChild, "PROJECT_?_SPF_SUBLINES", "0", "Info"
        .Add SPFNode, tvwChild, "PROJECT_?_SPF_FUNCLINES", "0", "Info"
        .Add SPFNode, tvwChild, "PROJECT_?_SPF_PROPLINES", "0", "Info"
        .Add SPFNode, tvwChild, "PROJECT_?_SPF_AVRSUBLINES", "0", "Info"
        .Add SPFNode, tvwChild, "PROJECT_?_SPF_AVRFUNCLINES", "0", "Info"
        .Add SPFNode, tvwChild, "PROJECT_?_SPF_AVRPROPLINES", "0", "Info"
    End With
End Sub

Private Sub AddDataToTreeview(PluralItemName As String, linedata As String, TVPic As String, Optional ResFile As String)    'Adds collected data about the current object to the treeview
    Dim LogFileDir As String, cpPos As Long, LoopVar As Long
    Dim Temp As Long, ParentItemNum As String, SPFData As String
    Dim TotalLines As Long

    On Local Error Resume Next

    With FrmResults.TreeView.Nodes ' Add calculated statistics to treeview
        ParentItemNum = "PROJECT_?_" & PluralItemName & "_" & ProjectItem.PrjItemName
        .Add GetNodeNum("PROJECT_?_" & PluralItemName), tvwChild, ParentItemNum, ProjectItem.PrjItemName, TVPic

        Temp = GetNodeNum(ParentItemNum)

        If TVPic <> "Module" And TVPic <> "Class" And TVPic <> "Designer" Then
            .Add Temp, tvwChild, ParentItemNum & "_CONTROLS", "控件 (包括所有数组对象): " & ProjectItem.PrjItemControls, "Info"
            .Add Temp, tvwChild, ParentItemNum & "_CONTROLSNA", "控件 (不包括所有数组对象): " & (ProjectItem.PrjItemControls - ProjectItem.PrjItemControlsNoArrays), "Info"
        End If

        TotalLines = ProjectItem.PrjItemHybridLines + ProjectItem.PrjItemCommentLines + ProjectItem.PrjItemCodeLines

        .Add Temp, tvwChild, ParentItemNum & "_LINES", "代码行数 (包括空行): " & ProjectItem.PrjItemCodeLines & " [" & Round((100 / IIf(TotalLines, TotalLines, 1)) * ProjectItem.PrjItemCodeLines, 2) & "%]", "Info"
        .Add Temp, tvwChild, ParentItemNum & "_LINESNB", "代码行数 (无空行): " & ProjectItem.PrjItemCodeLinesNoBlanks & " [" & Round((100 / IIf(TotalLines, TotalLines, 1)) * ProjectItem.PrjItemCodeLinesNoBlanks, 2) & "%]", "Info"
        .Add Temp, tvwChild, ParentItemNum & "_COMMENTLINES", "注释行数: " & ProjectItem.PrjItemCommentLines & " [" & Round((100 / IIf(TotalLines, TotalLines, 1)) * ProjectItem.PrjItemCommentLines, 2) & "%]", "Info"
        .Add Temp, tvwChild, ParentItemNum & "_HYBRIDLINES", "混合行数: " & ProjectItem.PrjItemHybridLines & " [" & Round((100 / IIf(TotalLines, TotalLines, 1)) * ProjectItem.PrjItemHybridLines, 2) & "%]", "Info"
        .Add Temp, tvwChild, ParentItemNum & "_VARIABLES", "声明变量: " & ProjectItem.PrjItemVariables, "Info"
        .Add Temp, tvwChild, ParentItemNum & "_CONSTANTS", "生命常数: " & ProjectItem.PrjItemConstants, "Info"
        .Add Temp, tvwChild, ParentItemNum & "_TYPES", "生命类型: " & ProjectItem.PrjItemTypes, "Info"
        .Add Temp, tvwChild, ParentItemNum & "_ENUMS", "声明枚举: " & ProjectItem.PrjItemEnums, "Info"

        .Add Temp, tvwChild, ParentItemNum & "_SUBS", "过程", "Method"
        .Add Temp, tvwChild, ParentItemNum & "_FUNCTIONS", "函数", "Method"
        .Add Temp, tvwChild, ParentItemNum & "_PROPERTIES", "属性", "Property"

        If TVPic <> "Module" And TVPic <> "Designer" Then .Add GetNodeNum(ParentItemNum), tvwChild, ParentItemNum & "_EVENTS", "事件", "Event"

        .Add Temp, tvwChild, ParentItemNum & "_STATMEMENTS", "语句", "CodeLoop"
        .Add GetNodeNum(ParentItemNum & "_STATMEMENTS"), tvwChild, ParentItemNum & "_STATMEMENTS_FOR", "For/Next: " & ProjectItem.PrjItemStatements(STFOR), "CodeLoop"
        .Add GetNodeNum(ParentItemNum & "_STATMEMENTS"), tvwChild, ParentItemNum & "_STATMEMENTS_DO", "Do/Loop: " & ProjectItem.PrjItemStatements(STDO), "CodeLoop"
        .Add GetNodeNum(ParentItemNum & "_STATMEMENTS"), tvwChild, ParentItemNum & "_STATMEMENTS_WHILE", "While/Wend: " & ProjectItem.PrjItemStatements(STWHILE), "CodeLoop"
        .Add GetNodeNum(ParentItemNum & "_STATMEMENTS"), tvwChild, ParentItemNum & "_STATMEMENTS_IF", "If/End If: " & ProjectItem.PrjItemStatements(STIF), "CodeLoop"
        .Add GetNodeNum(ParentItemNum & "_STATMEMENTS"), tvwChild, ParentItemNum & "_STATMEMENTS_SELECT", "Select/End Select: " & ProjectItem.PrjItemStatements(STSELECT), "CodeLoop"

        If CheckForMalicious Then
            .Add Temp, tvwChild, ParentItemNum & "_PMC", "潜在恶意代码", "BadCode"
            .Add GetNodeNum(ParentItemNum & "_PMC"), tvwChild, ParentItemNum & "_PMC_COUNT", "总计: " & UBound(PMaliciousCode), "Info"

            If UBound(PMaliciousCode) Then
                BubbleSortArray PMaliciousCode ' Sort the potentially malicious code array into alphabetical order

                For LoopVar = 1 To UBound(PMaliciousCode)
                    ' This SHOULD be 1 - first item (index 0) is always blank due to design
                    .Add GetNodeNum(ParentItemNum & "_PMC"), tvwChild, ParentItemNum & "_PMC_ITEM" & LoopVar, PMaliciousCode(LoopVar), "BadCode"
                Next
            End If

            ReDim PMaliciousCode(0) As String ' Clear the array
        End If

        .Item(GetNodeNum("PROJECT_?_" & PluralItemName & "_LINES")).Text = Int(.Item(GetNodeNum("PROJECT_?_" & PluralItemName & "_LINES")).Text) + ProjectItem.PrjItemCodeLines
        .Item(GetNodeNum("PROJECT_?_" & PluralItemName & "_VARIABLES")).Text = Int(.Item(GetNodeNum("PROJECT_?_" & PluralItemName & "_VARIABLES")).Text) + ProjectItem.PrjItemVariables

        Project.ProjectLines = ProjectItem.PrjItemCodeLines + ProjectItem.PrjItemHybridLines ' Add to the project total lines (hybrid lines are NOT included as code lines, so must be added separately)
        Project.ProjectLinesNB = ProjectItem.PrjItemCodeLinesNoBlanks + ProjectItem.PrjItemHybridLines ' Ditto ;)

        If GetSetting("DeepLook", "Options", "SortSPFs", 1) Then ProjectItem.SortArrays ' If option set, alphabetically sort SPF names

        Temp = ProjectItem.PrjItemItemSubsCount
        .Add GetNodeNum(ParentItemNum & "_SUBS"), tvwChild, ParentItemNum & "_SUB_COUNT", "总计: " & Temp, "Info"
        For LoopVar = 1 To Temp
            SPFData = ProjectItem.PrjItemItemSubs(LoopVar)
            .Add GetNodeNum(ParentItemNum & "_SUBS"), tvwChild, ParentItemNum & "_SUB" & LoopVar, Left$(SPFData, InStr(1, SPFData, ";") - 1), "Method"
            If ShowIndividualLines Then
                cpPos = InStr(1, SPFData, ";") + 1
                .Add GetNodeNum(ParentItemNum & "_SUB" & LoopVar), tvwChild, ParentItemNum & "_SUB" & LoopVar & "_LINES", "代码行数 (包含空行): " & Mid$(SPFData, cpPos, InStrRev(SPFData, ":") - cpPos), "Info"
                .Add GetNodeNum(ParentItemNum & "_SUB" & LoopVar), tvwChild, ParentItemNum & "_SUB" & LoopVar & "_LINESNB", "代码行数 (无空行): " & Mid$(SPFData, InStr(1, SPFData, ":") + 1), "Info"
            End If
        Next

        Temp = ProjectItem.PrjItemItemFunctionsCount
        .Add GetNodeNum(ParentItemNum & "_FUNCTIONS"), tvwChild, ParentItemNum & "_FUNCTION_COUNT", "总计: " & Temp, "Info"
        For LoopVar = 1 To Temp
            SPFData = ProjectItem.PrjItemItemFunctions(LoopVar)
            .Add GetNodeNum(ParentItemNum & "_FUNCTIONS"), tvwChild, ParentItemNum & "_FUNCTION" & LoopVar, Left$(SPFData, InStr(1, SPFData, ";") - 1), "Method"
            If ShowIndividualLines Then
                cpPos = InStr(1, SPFData, ";") + 1
                .Add GetNodeNum(ParentItemNum & "_FUNCTION" & LoopVar), tvwChild, ParentItemNum & "_FUNCTION" & LoopVar & "_LINES", "代码行数 (包含空行): " & Mid$(SPFData, cpPos, InStrRev(SPFData, ":") - cpPos), "Info"
                .Add GetNodeNum(ParentItemNum & "_FUNCTION" & LoopVar), tvwChild, ParentItemNum & "_FUNCTION" & LoopVar & "_LINESNB", "代码行数 (无空行):  " & Mid$(SPFData, InStr(1, SPFData, ":") + 1), "Info"
            End If
        Next

        Temp = ProjectItem.PrjItemItemPropertiesCount
        .Add GetNodeNum(ParentItemNum & "_PROPERTIES"), tvwChild, ParentItemNum & "_PROPERTY_COUNT", "总计: " & Temp, "Info"
        For LoopVar = 1 To Temp
            SPFData = ProjectItem.PrjItemItemProperties(LoopVar)
            .Add GetNodeNum(ParentItemNum & "_PROPERTIES"), tvwChild, ParentItemNum & "_PROPERTY" & LoopVar, Left$(SPFData, InStr(1, SPFData, ";") - 1), "Property"
            If ShowIndividualLines Then
                cpPos = InStr(1, ProjectItem.PrjItemItemProperties(LoopVar), ";") + 1
                .Add GetNodeNum(ParentItemNum & "_PROPERTY" & LoopVar), tvwChild, ParentItemNum & "_PROPERTY" & LoopVar & "_LINES", "代码行数 (包含空行): " & Mid$(SPFData, cpPos, InStrRev(SPFData, ":") - cpPos), "Info"
                .Add GetNodeNum(ParentItemNum & "_PROPERTY" & LoopVar), tvwChild, ParentItemNum & "_PROPERTY" & LoopVar & "_LINESNB", "代码行数 (无空行): " & Mid$(SPFData, InStr(1, SPFData, ":") + 1), "Info"
            End If
        Next

        If TVPic <> "Module" And TVPic <> "Designer" Then
            Temp = GetNodeNum(ParentItemNum & "_EVENTS")
            .Add Temp, tvwChild, ParentItemNum & "_EVENTS_COUNT", "总计: " & ProjectItem.PrjItemItemEventsCount, "Info"
            For LoopVar = 1 To ProjectItem.PrjItemItemEventsCount
                .Add Temp, tvwChild, ParentItemNum & "_EVENT" & LoopVar, ProjectItem.PrjItemItemEvents(LoopVar), "Event"
            Next
        End If

        Temp = GetNodeNum(ParentItemNum)
        If LenB(ResFile) Then
            .Add Temp, tvwChild, ParentItemNum & "_" & ResFile & "FILE", ResFile & " 资源文件", "RelDoc"
            FileInfo.FindFileInfo GetRootDirectory(Project.ProjectPath) & Left$(linedata, Len(linedata) - 3) & ResFile, False

            If FileInfo.ByteSize <> "bytes" Then
                .Add GetNodeNum(ParentItemNum & "_" & ResFile & "FILE"), tvwChild, ParentItemNum & "_" & ResFile & "FILE_FILESIZE", "文件大小: " & FileInfo.ByteSize, "Info"
            Else
                .Add GetNodeNum(ParentItemNum & "_" & ResFile & "FILE"), tvwChild, ParentItemNum & "_" & ResFile & "FILE_FILESIZE", "文件大小: N/A", "Info"
            End If
        End If

        .Add Temp, tvwChild, ParentItemNum & "_FILESTATS", "文件信息", "LOGFile"

        Temp = GetNodeNum(ParentItemNum & "_FILESTATS")
        FileInfo.FindFileInfo GetRootDirectory(Project.ProjectPath) & linedata, False
        .Add Temp, tvwChild, ParentItemNum & "_FILESTATS_FILEMODIFIED", "最近修改时间: " & FileInfo.LastWriteTime, "Info"
        .Add Temp, tvwChild, ParentItemNum & "_FILESTATS_FILEOPENED", "最近访问时间: " & FileInfo.LastAccessTime, "Info"
        .Add Temp, tvwChild, ParentItemNum & "_FILESTATS_FILESIZE", "文件大小: " & FileInfo.ByteSize, "Info"

        LogFileDir = GetRootDirectory(Project.ProjectPath) & (Left$(linedata, Len(linedata) - 4)) & ".log"
        If FileExists(LogFileDir) Then ' Check for a log file
            Temp = GetNodeNum(ParentItemNum)
            .Add Temp, tvwChild, ParentItemNum & "_LOGFILE_" & LogFileDir, "(双击查看日志文件)", "LOGFile"
        End If
    End With
End Sub

Private Sub FixPrjTreeViewItems() ' Sub to fix up captions and keys of the treeview's nodes
    Dim ItemKeyNum As Long, QuestionPos As Long, OnePercent As Single, SPFColourIndex As String * 3
    Dim TempStr As String

    CurrentScanFile "Clean" ' Put the "clean" symbol on the small picture showing the current action

    With FrmSelProject
        .pgbAPB.Value = 0 ' Reset the progressbar
        .pgbAPB.Color = 6340579 ' Set the progressbar to a yellow colour (treeview cleaning phase)
        .lblScanPhase.Caption = "Treeview Cleanup Phase" ' Change phase caption
        .lblScanPhase.ForeColor = 6340579 ' Set the colour to that of the progressbar
    End With

    DoEvents

    With FrmResults.TreeView.Nodes ' "With" statement looks better - cleaner code
        ItemKeyNum = GetNodeNum("PROJECT_?")
        .Item(ItemKeyNum).Text = Project.ProjectName

        ItemKeyNum = GetNodeNum("PROJECT_?_TITLE")
        .Item(ItemKeyNum).Text = "标题: " & IIf(Project.ProjectTitle = vbNullString, Project.ProjectName, Project.ProjectTitle)

        ItemKeyNum = GetNodeNum("PROJECT_?_VERSION")
        .Item(ItemKeyNum).Text = "版本: " & Project.ProjectVersion

        ItemKeyNum = GetNodeNum("PROJECT_?_TYPE")
        .Item(ItemKeyNum).Text = "类型: " & Project.ProjectProjectType

        ItemKeyNum = GetNodeNum("PROJECT_?_STARTUPITEM")
        .Item(ItemKeyNum).Text = "启动对象: " & Project.ProjectStartupItem

        ItemKeyNum = GetNodeNum("PROJECT_?_LINES")
        .Item(ItemKeyNum).Text = "代码行数 (包含空行): " & IIf(Project.ProjectLines, Format$(Project.ProjectLines, "###,###,###") & " [" & Round((100 / (Project.ProjectLines + Project.ProjectCommentLines)) * Project.ProjectLines, 2) & "%]", "0")

        TotalLines = TotalLines + Project.ProjectLines

        ItemKeyNum = GetNodeNum("PROJECT_?_LINESNB")
        .Item(ItemKeyNum).Text = "代码行数 (无空行): " & IIf(Project.ProjectLinesNB, Format$(Project.ProjectLinesNB, "###,###,###") & " [" & Round((100 / (Project.ProjectLines + Project.ProjectCommentLines)) * Project.ProjectLinesNB, 2) & "%]", "0")

        ItemKeyNum = GetNodeNum("PROJECT_?_LINESCOMMENT")
        .Item(ItemKeyNum).Text = "代码行数 (注释): " & IIf(Project.ProjectCommentLines, Format$(Project.ProjectCommentLines, "###,###,###") & " [" & Round((100 / (Project.ProjectLines + Project.ProjectCommentLines)) * Project.ProjectCommentLines, 2) & "%]", "0")

        ItemKeyNum = GetNodeNum("PROJECT_?_FORMS_COUNT")
        .Item(ItemKeyNum).Text = "总计: " & Project.ProjectForms
        ItemKeyNum = GetNodeNum("PROJECT_?_MODULES_COUNT")
        .Item(ItemKeyNum).Text = "总计: " & Project.ProjectModules
        ItemKeyNum = GetNodeNum("PROJECT_?_CLASSES_COUNT")
        .Item(ItemKeyNum).Text = "总计: " & Project.ProjectClasses
        ItemKeyNum = GetNodeNum("PROJECT_?_USERCONTROLS_COUNT")
        .Item(ItemKeyNum).Text = "总计: " & Project.ProjectUserControls
        ItemKeyNum = GetNodeNum("PROJECT_?_USERDOCUMENTS_COUNT")
        .Item(ItemKeyNum).Text = "总计: " & Project.ProjectUserDocuments
        ItemKeyNum = GetNodeNum("PROJECT_?_PROPERTYPAGES_COUNT")
        .Item(ItemKeyNum).Text = "总计: " & Project.ProjectPropertyPages
        ItemKeyNum = GetNodeNum("PROJECT_?_DESIGNERS_COUNT")
        .Item(ItemKeyNum).Text = "总计: " & Project.ProjectDesigners

        ItemKeyNum = GetNodeNum("PROJECT_?_FORMS_LINES")
        .Item(ItemKeyNum).Text = "代码行数: " & MaskData(.Item(ItemKeyNum).Text)

        ItemKeyNum = GetNodeNum("PROJECT_?_MODULES_LINES")
        .Item(ItemKeyNum).Text = "代码行数: " & MaskData(.Item(ItemKeyNum).Text)

        ItemKeyNum = GetNodeNum("PROJECT_?_CLASSES_LINES")
        .Item(ItemKeyNum).Text = "代码行数: " & MaskData(.Item(ItemKeyNum).Text)

        ItemKeyNum = GetNodeNum("PROJECT_?_USERCONTROLS_LINES")
        .Item(ItemKeyNum).Text = "代码行数: " & MaskData(.Item(ItemKeyNum).Text)

        ItemKeyNum = GetNodeNum("PROJECT_?_USERDOCUMENTS_LINES")
        .Item(ItemKeyNum).Text = "代码行数: " & MaskData(.Item(ItemKeyNum).Text)

        ItemKeyNum = GetNodeNum("PROJECT_?_PROPERTYPAGES_LINES")
        .Item(ItemKeyNum).Text = "代码行数: " & MaskData(.Item(ItemKeyNum).Text)

        ItemKeyNum = GetNodeNum("PROJECT_?_DESIGNERS_LINES")
        .Item(ItemKeyNum).Text = "代码行数: " & MaskData(.Item(ItemKeyNum).Text)

        ItemKeyNum = GetNodeNum("PROJECT_?_FORMS_VARIABLES")
        .Item(ItemKeyNum).Text = "声明变量: " & MaskData(.Item(ItemKeyNum).Text)

        ItemKeyNum = GetNodeNum("PROJECT_?_MODULES_VARIABLES")
        .Item(ItemKeyNum).Text = "声明变量: " & MaskData(.Item(ItemKeyNum).Text)

        ItemKeyNum = GetNodeNum("PROJECT_?_CLASSES_VARIABLES")
        .Item(ItemKeyNum).Text = "声明变量: " & MaskData(.Item(ItemKeyNum).Text)

        ItemKeyNum = GetNodeNum("PROJECT_?_USERCONTROLS_VARIABLES")
        .Item(ItemKeyNum).Text = "声明变量: " & MaskData(.Item(ItemKeyNum).Text)

        ItemKeyNum = GetNodeNum("PROJECT_?_USERDOCUMENTS_VARIABLES")
        .Item(ItemKeyNum).Text = "声明变量: " & MaskData(.Item(ItemKeyNum).Text)

        ItemKeyNum = GetNodeNum("PROJECT_?_PROPERTYPAGES_VARIABLES")
        .Item(ItemKeyNum).Text = "声明变量: " & MaskData(.Item(ItemKeyNum).Text)

        ItemKeyNum = GetNodeNum("PROJECT_?_DESIGNERS_VARIABLES")
        .Item(ItemKeyNum).Text = "声明变量: " & MaskData(.Item(ItemKeyNum).Text)

        ItemKeyNum = GetNodeNum("PROJECT_?_VARIABLES")
        .Item(ItemKeyNum).Text = "声明变量: " & ZeroIfNull(Format$(Project.ProjectVariables, "###,###,###"))

        ItemKeyNum = GetNodeNum("PROJECT_?_CONSTANTS")
        .Item(ItemKeyNum).Text = "声明常数: " & ZeroIfNull(Format$(Project.ProjectConstants, "###,###,###"))

        ItemKeyNum = GetNodeNum("PROJECT_?_TYPES")
        .Item(ItemKeyNum).Text = "声明类型: " & ZeroIfNull(Format$(Project.ProjectTypes, "###,###,###"))

        ItemKeyNum = GetNodeNum("PROJECT_?_ENUMS")
        .Item(ItemKeyNum).Text = "声明枚举: " & ZeroIfNull(Format$(Project.ProjectEnums, "###,###,###"))

        ItemKeyNum = GetNodeNum("PROJECT_?_REFCOM_COUNT")
        .Item(ItemKeyNum).Text = "总计: " & Project.ProjectRefComCount

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_SUBS")
        TotalSubs = Int(.Item(ItemKeyNum).Text)
        .Item(ItemKeyNum).Text = "合计子程序: " & .Item(ItemKeyNum).Text

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_FUNCTIONS")
        TotalFunctions = Int(.Item(ItemKeyNum).Text)
        .Item(ItemKeyNum).Text = "合计函数: " & ZeroIfNull(Format$(Int(.Item(ItemKeyNum).Text), "###,###,###"))

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_PROPERTIES")
        TotalProperties = Int(.Item(ItemKeyNum).Text)
        .Item(ItemKeyNum).Text = "合计属性: " & ZeroIfNull(Format$(Int(.Item(ItemKeyNum).Text), "###,###,###"))

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_EVENTS")
        TotalEvents = Int(.Item(ItemKeyNum).Text)
        .Item(ItemKeyNum).Text = "合计事件: " & ZeroIfNull(Format$(Int(TotalEvents), "###,###,###"))

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_DECLAREDSUBS")
        TotalDecSubs = Int(.Item(ItemKeyNum).Text)
        .Item(ItemKeyNum).Text = "过程声明合计: " & ZeroIfNull(Format$(Int(.Item(ItemKeyNum).Text), "###,###,###"))

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_DECLAREDFUNCTIONS")
        TotalDecFunctions = Int(.Item(ItemKeyNum).Text)
        .Item(ItemKeyNum).Text = "合计声明函数: " & ZeroIfNull(Format$(Int(.Item(ItemKeyNum).Text), "###,###,###"))

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_SUBLINES")
        .Item(ItemKeyNum).Text = "过程代码行数: " & ZeroIfNull(Format$(Project.ProjectSubLines, "###,###,###"))

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_FUNCLINES")
        .Item(ItemKeyNum).Text = "函数代码函数: " & ZeroIfNull(Format$(Project.ProjectFuncLines, "###,###,###"))

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_PROPLINES")
        .Item(ItemKeyNum).Text = "属性代码函数: " & ZeroIfNull(Format$(Project.ProjectPropLines, "###,###,###"))

        ItemKeyNum = GetNodeNum("PROJECT_?_RELATEDDOCUMENTS_COUNT")
        .Item(ItemKeyNum).Text = "总计: " & TotalRelDocs

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_AVRSUBLINES")
        .Item(ItemKeyNum).Text = "过程平均代码行数: " & Format$(Round(Project.ProjectSubLines / IIf(TotalSubs, TotalSubs, 1), 2), "###,###.##")

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_AVRFUNCLINES")
        .Item(ItemKeyNum).Text = "函数平均代码行数: " & Format$(Round(Project.ProjectFuncLines / IIf(TotalFunctions, TotalFunctions, 1), 2), "###,###.##")

        ItemKeyNum = GetNodeNum("PROJECT_?_SPF_AVRPROPLINES")
        .Item(ItemKeyNum).Text = "属性平均代码行数: " & Format$(Round(Project.ProjectPropLines / IIf(TotalProperties, TotalProperties, 1), 2), "###,###.##")

        OnePercent = (100 / (.Count - ThisPrgNodeStart)) ' Calculate what one percent completion of the current task is

        If UsesGroup Then ' This is made faster by separating the code blocks, rather than evaluation an "If UsesGroup = True" every time
            For ItemKeyNum = ThisPrgNodeStart To .Count ' Go over every node, checking for Related Documents, and replacing the temp project name "?" with the correct
                .Item(ItemKeyNum).Key = Replace$(.Item(ItemKeyNum).Key, "?", Project.ProjectName, 8, 1)

                Do
                    QuestionPos = InStr(1, .Item(ItemKeyNum).Text, "  ") ' Checks for indents, which make line continuations look wrong
                    If QuestionPos Then  ' "QuestionPos" variable used here only to save on variable requirements
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, QuestionPos - 1) & Mid$(.Item(ItemKeyNum).Text, QuestionPos + 2)
                    Else ' No indents left - exit the loop
                        Exit Do
                    End If
                Loop

                If Int(ItemKeyNum * 0.01) = ItemKeyNum * 0.01 Then FrmSelProject.pgbAPB.Value = OnePercent * (ItemKeyNum - ThisPrgNodeStart) ' Only update the progressbar every 100 nodes to save time

                If InStrB(1, .Item(ItemKeyNum).Text, "[") Then
                    SPFColourIndex = Right$(.Item(ItemKeyNum).Text, 3) ' Save time by getting this once per loop

                    ' The colour numbers are out of order because the most frequently used types should be tested before obscure (like static or friend subs) SPF's.
                    If InStrB(1, .Item(ItemKeyNum).Text, "[EXT]") Then  'Is an external (DLL) call (declared sub or function)
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(150, 150, 150) 'Make key text a light-grey colour
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 6) 'Remove the now unnecessary "[EXT]" text
                    ElseIf SPFColourIndex = "[1]" Then 'SPF Colour 1
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(130, 0, 200) 'Colour Key Text
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                    ElseIf SPFColourIndex = "[2]" Then 'SPF Colour 2
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(200, 0, 150) 'Colour Key Text
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                    ElseIf SPFColourIndex = "[3]" Then 'SPF Colour 3
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(10, 150, 10) 'Colour Key Text
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                    ElseIf SPFColourIndex = "[6]" Then 'SPF Colour 4
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(249, 164, 0) 'Colour Key Text
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                    ElseIf SPFColourIndex = "[7]" Then 'SPF Colour 5
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(217, 206, 19) 'Colour Key Text
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                    ElseIf SPFColourIndex = "[8]" Then 'SPF Colour 6
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(19, 217, 192) 'Colour Key Text
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                    ElseIf SPFColourIndex = "[4]" Then 'SPF Colour 7
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(20, 90, 100) 'Colour Key Text
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                    ElseIf SPFColourIndex = "[5]" Then 'SPF Colour 8
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(50, 23, 80) 'Colour Key Text
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                    End If
                End If

                If InStrB(1, .Item(ItemKeyNum).Text, "无空格):") Then   'No Blanks
                    QuestionPos = InStr(1, .Item(ItemKeyNum).Text, ":") + 2

                    TempStr = Mid$(.Item(ItemKeyNum).Text, QuestionPos)
                    If InStr(QuestionPos, .Item(ItemKeyNum).Text, " ") Then
                        TempStr = Left$(TempStr, InStr(1, TempStr, " "))
                    End If

                    If TempStr <> "N/A" Then ' Not a declared sub/function
                        If Int(TempStr) = 0 Then ' Empty SPF
                            If InStrB(1, .Item(ItemKeyNum).Parent.Key, "_SUB") And InStrB(1, .Item(ItemKeyNum).Parent.Text, " Lib ") = 0 Then
                                .Item(ItemKeyNum).Parent.ForeColor = vbRed
                            ElseIf InStrB(1, .Item(ItemKeyNum).Parent.Key, "_FUNC") And InStrB(1, .Item(ItemKeyNum).Parent.Text, " Lib ") = 0 Then
                                .Item(ItemKeyNum).Parent.ForeColor = vbRed
                            ElseIf InStrB(1, .Item(ItemKeyNum).Parent.Key, "_PROP") Then
                                .Item(ItemKeyNum).Parent.ForeColor = vbRed
                            End If
                        End If
                    End If
                End If
            Next
        Else 'Identical to the loop for group projects, but doesn't need to replace the "?" in the node keys
            For ItemKeyNum = 1 To .Count 'Go over every node, checking for Related Documents, and replacing the temp project name "?" with the correct
                Do
                    QuestionPos = InStr(1, .Item(ItemKeyNum).Text, "  ") ' Checks for indents, which make line continuations look wrong
                    If QuestionPos Then  ' "QuestionPos" variable used here only to save on variable requirements
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, QuestionPos - 1) & Mid$(.Item(ItemKeyNum).Text, QuestionPos + 2)
                    Else ' No indents left - exit the loop
                        Exit Do
                    End If
                Loop

                If Int(ItemKeyNum * 0.02) = ItemKeyNum * 0.02 Then FrmSelProject.pgbAPB.Value = OnePercent * ItemKeyNum ' Save time by only updating every 50 nodes

                If InStrB(1, .Item(ItemKeyNum).Text, "[") Then
                    SPFColourIndex = Right$(.Item(ItemKeyNum).Text, 3) ' Save time by getting this once per loop

                    ' The colour numbers are out of order because the most frequently used types should be tested before obscure (like static or friend subs) SPF's.
                    If InStrB(1, .Item(ItemKeyNum).Text, "[EXT]") Then  'Is an external (DLL) call (declared sub or function)
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(150, 150, 150) 'Make key text a light-grey colour
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 6) 'Remove the now unnecessary "[EXT]" text
                    ElseIf SPFColourIndex = "[1]" Then 'SPF Colour 1
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(130, 0, 200) 'Colour Key Text
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                    ElseIf SPFColourIndex = "[2]" Then 'SPF Colour 2
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(200, 0, 150) 'Colour Key Text
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                    ElseIf SPFColourIndex = "[3]" Then 'SPF Colour 3
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(10, 150, 10) 'Colour Key Text
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                    ElseIf SPFColourIndex = "[6]" Then 'SPF Colour 4
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(249, 164, 0) 'Colour Key Text
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                    ElseIf SPFColourIndex = "[7]" Then 'SPF Colour 5
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(217, 206, 19) 'Colour Key Text
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                    ElseIf SPFColourIndex = "[8]" Then 'SPF Colour 6
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(19, 217, 192) 'Colour Key Text
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                    ElseIf SPFColourIndex = "[4]" Then 'SPF Colour 7
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(20, 90, 100) 'Colour Key Text
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                    ElseIf SPFColourIndex = "[5]" Then 'SPF Colour 8
                        .Item(ItemKeyNum).Parent.ForeColor = RGB(50, 23, 80) 'Colour Key Text
                        .Item(ItemKeyNum).Text = Left$(.Item(ItemKeyNum).Text, Len(.Item(ItemKeyNum).Text) - 3) 'Remove the now unnecessary "[x]" text
                    End If
                End If

                If InStrB(1, .Item(ItemKeyNum).Text, "无空格):") Then  'No Blanks
                    QuestionPos = InStr(1, .Item(ItemKeyNum).Text, ":") + 2

                    TempStr = Mid$(.Item(ItemKeyNum).Text, QuestionPos)
                    If InStr(QuestionPos, .Item(ItemKeyNum).Text, " ") Then
                        TempStr = Left$(TempStr, InStr(1, TempStr, " "))
                    End If

                    If TempStr <> "N/A" Then ' Not a declared sub/function
                        If Int(TempStr) = 0 Then ' Empty SPF
                            If InStrB(1, .Item(ItemKeyNum).Parent.Key, "_SUB") And InStrB(1, .Item(ItemKeyNum).Parent.Text, " Lib ") = 0 Then
                                .Item(ItemKeyNum).Parent.ForeColor = vbRed
                            ElseIf InStrB(1, .Item(ItemKeyNum).Parent.Key, "_FUNC") And InStrB(1, .Item(ItemKeyNum).Parent.Text, " Lib ") = 0 Then
                                .Item(ItemKeyNum).Parent.ForeColor = vbRed
                            ElseIf InStrB(1, .Item(ItemKeyNum).Parent.Key, "_PROP") Then
                                .Item(ItemKeyNum).Parent.ForeColor = vbRed
                            End If
                        End If
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub GetProjectEXEStats()
    Dim EXEFileVer() As String, PROJFileVer() As String

    EXENewOrOld = "NF"

    If InStr(1, Project.ProjectEXEPath, ":") Then ' Full path specified in the EXE path variable
        If FileExists(Project.ProjectEXEPath & Project.ProjectEXEFName) = False Or LenB(Project.ProjectEXEFName) = 0 Then Exit Sub

        FileInfo.FindFileInfo Project.ProjectEXEPath & Project.ProjectEXEFName, False
    Else
        If FileExists(GetRootDirectory(Project.ProjectPath) & Project.ProjectEXEPath & Project.ProjectEXEFName) = False Or LenB(Project.ProjectEXEFName) = 0 Then Exit Sub

        FileInfo.FindFileInfo GetRootDirectory(Project.ProjectPath) & Project.ProjectEXEPath & Project.ProjectEXEFName, False
    End If

    With FrmResults.TreeView.Nodes
        .Add GetNodeNum("PROJECT_?"), tvwChild, "PROJECT_?_EXESTATS", "可执行文件", "App"
        .Add GetNodeNum("PROJECT_?_EXESTATS"), tvwChild, "PROJECT_?_EXESTATS_NAME", "文件名称: " & Project.ProjectEXEFName, "Info"
        .Add GetNodeNum("PROJECT_?_EXESTATS"), tvwChild, "PROJECT_?_EXESTATS_PATH", "文件路径: " & FixRelPath(GetRootDirectory(Project.ProjectPath) & Project.ProjectEXEPath), "Info"
        .Add GetNodeNum("PROJECT_?_EXESTATS"), tvwChild, "PROJECT_?_EXESTATS_SIZE", "文件大小: " & FileInfo.ByteSize, "Info"
        .Add GetNodeNum("PROJECT_?_EXESTATS"), tvwChild, "PROJECT_?_EXESTATS_CTIME", "创建时间: " & FileInfo.CreationTime, "Info"
        .Add GetNodeNum("PROJECT_?_EXESTATS"), tvwChild, "PROJECT_?_EXESTATS_VERSION", "文件版本: " & FileInfo.FileVersion, "Info"
    End With

    EXEFileVer = Split(FileInfo.FileVersion, ".")
    PROJFileVer = Split(Project.ProjectVersion, ".")

    EXENewOrOld = CompareVersions(PROJFileVer, EXEFileVer)
End Sub

Private Sub GetRelatedDocStats(FileName As String)  ' No stats here, just adding the file to the treeview
    Dim FixedFileName As String

    If InStrB(1, FileName, ".RES", vbTextCompare) Then
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_RELATEDDOCUMENTS"), tvwChild, "PROJECT_?_RELATEDDOCUMENT_" & FileName, FileName, "Resource"
    Else
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_RELATEDDOCUMENTS"), tvwChild, "PROJECT_?_RELATEDDOCUMENT_" & FileName, FileName, "RelDoc"
    End If

    TotalRelDocs = TotalRelDocs + 1 ' Increase the Related Document count, and add to a string for the report
    RelatedDocumentFNames = RelatedDocumentFNames & FileName & Space$(59 - Len(FileName)) & " |" & vbNewLine

    FixedFileName = FileName

    If Mid$(FileName, 2, 1) <> ":" Then
        If Not FileExists(GetRootDirectory(Project.ProjectPath) & FileName) Then
            If FileExists(FilesRootDirectory & GetRootDirectory(Project.ProjectPath) & FileName) Then
                FixedFileName = FilesRootDirectory & GetRootDirectory(Project.ProjectPath) & FileName
            End If
        Else
            FixedFileName = GetRootDirectory(Project.ProjectPath) & FileName
        End If
    End If

    FileInfo.FindFileInfo FixedFileName, False

    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_RELATEDDOCUMENT_" & FileName), tvwChild, "PROJECT_?_RELATEDDOCUMENT_" & FileName & "_SIZE", "File Size: " & FileInfo.ByteSize, "Info"
End Sub

Private Sub GetComponentStats(FileName As String)  ' Get statistics on the component passed
    Dim ClassName As String, OCXFileName As String, ColonPos As Long, HashPos As Long, SecondHashPos As Long

    If InStrB(1, FileName, ".vbp", vbTextCompare) Then GoTo IsProjectComponent ' The component is actually a project in a group, skip to the appropriate section

    Project.ProjectRefComCount = 1 ' The class actually adds one, rather than setting it as one - saves code

    ColonPos = InStr(1, FileName, ";") ' Finds the colon in the string separating the name and GUID (class name)
    OCXFileName = Mid$(FileName, ColonPos + 1)
    ClassName = Left$(FileName, ColonPos - 1)

    HashPos = InStr(1, ClassName, "#")
    If HashPos = 0 Then GoTo InsertableComponent ' Insertable Components don't show a GUID in the project file
    SecondHashPos = InStrRev(ClassName, "#") - 2

    ' Add component and class name to treeview:
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM"), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & OCXFileName, OCXFileName, "Component"
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM_COMPONENT_" & OCXFileName), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & OCXFileName & "_CLASSNAME", "Class Name: " & Left$(ClassName, HashPos - 1), "Info"

    ClassName = Left$(ClassName, HashPos - 1) & "\" & Mid$(ClassName, HashPos + 1, Len(ClassName) - SecondHashPos)

    FileInfo.FindFileInfo GetComponentNameFromReg(ClassName), False

    ' Add component file info to treeview:
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM_COMPONENT_" & OCXFileName), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & OCXFileName & "_PATH", "文件路径: " & FileInfo.FileName, "Info"
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM_COMPONENT_" & OCXFileName), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & OCXFileName & "_DESCRIPTION", "描述: " & GetComponentDescFromReg(ClassName), "Info"

    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM_COMPONENT_" & OCXFileName), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & OCXFileName & "_FILEINFO", "文件信息", "LOGFile"
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM_COMPONENT_" & OCXFileName & "_FILEINFO"), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & OCXFileName & "_FILEINFO_SIZE", "文件大小: " & FileInfo.ByteSize, "Info"
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM_COMPONENT_" & OCXFileName & "_FILEINFO"), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & OCXFileName & "_FILEINFO_COMPANY", "公司名称: " & FileInfo.CompanyName, "Info"
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM_COMPONENT_" & OCXFileName & "_FILEINFO"), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & OCXFileName & "_FILEINFO_VERSION", "版本: " & FileInfo.FileVersion, "Info"

    ' Add component name and description to an array to make report creation easier:
    Project.AddData SPF_RefCom, GetComponentNameFromReg(ClassName)
    Project.AddData SPF_RefCom, GetComponentDescFromReg(ClassName)
    Exit Sub

IsProjectComponent:                                                                                                                                                                                                                                                                                                                         'Add the project component to the treeview
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM"), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & ExtractFileName(FileName), ExtractFileName(FileName), "Project"
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM_COMPONENT_" & ExtractFileName(FileName)), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & ExtractFileName(FileName) & "_PATH", "Project Path: " & FileName, "Info"

    Project.AddData SPF_RefCom, ExtractFileName(FileName)  ' Add filename to array
    Project.AddData SPF_RefCom, "(Project)"  ' Add a unknown description to array
    Exit Sub

InsertableComponent:
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM"), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & Left$(FileName, ColonPos - 1), Left$(FileName, ColonPos - 1), "IComponent"
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM_COMPONENT_" & Left$(FileName, ColonPos - 1)), tvwChild, "PROJECT_?_REFCOM_COMPONENT_" & Left$(FileName, ColonPos - 1) & "_PROGRAM", "Parent Program: " & Mid$(FileName, ColonPos + 1), "Info"
End Sub

Private Sub GetReferenceStats(linedata As String)  ' Get information on a reference (.dll, etc.)
    Dim cPos As Long, RefName As String, RefDesc As String, FirstHash As Long

    On Local Error Resume Next

    If LCase$(Right$(linedata, 4)) = ".vbp" Then GoTo IsProjectReference

    cPos = InStrRev(linedata, "#")
    RefDesc = Mid$(linedata, cPos + 1)

    FirstHash = cPos
    cPos = InStrRev(linedata, "#", FirstHash - 1)

    Project.ProjectRefComCount = 1 ' This actually adds one - saves code by writing the arithmetic once in the class

    RefName = ExtractFileName(Mid$(linedata, cPos + 1, (FirstHash - cPos) - 1))

    ' Add reference to treeview:
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM"), tvwChild, "PROJECT_?_REFCOM_REFERENCE_" & RefName, RefName, IsSysDLL(RefName)
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM_REFERENCE_" & RefName), tvwChild, "PROJECT_?_REFCOM_REFERENCE_" & RefName & "_DESC", "描述: " & RefDesc, "Info"

    ' Add reference to array:
    Project.AddData SPF_RefCom, RefName
    Project.AddData SPF_RefCom, RefDesc

    Exit Sub

IsProjectReference:
    If InStrRev(linedata, "\") Then
        RefName = Mid$(linedata, InStrRev(linedata, "\") + 1)
    Else
        RefName = linedata
    End If

    RefDesc = "Referenced Project"

    Project.AddData SPF_RefCom, RefName
    Project.AddData SPF_RefCom, RefDesc

    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM"), tvwChild, "PROJECT_?_REFCOM_REFERENCE_" & RefName, RefName, "Project"
    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_REFCOM_REFERENCE_" & RefName), tvwChild, "PROJECT_?_REFCOM_REFERENCE_" & RefName & "_DESC", "描述: 工程引用", "Info"
End Sub

Private Sub GetDecDllStats(linedata As String)  ' Gets statistics on a declared DLL
    Dim StartQ As Long, EndQ As Long, DLLFName As String
    Dim DLLRootNodeNum As Long

    StartQ = InStr(1, linedata, """") ' Find start of filename
    EndQ = InStr(StartQ + 1, linedata, """") ' Find end of filename

    If StartQ = 0 Or EndQ = 0 Then Exit Sub

    DLLFName = Mid$(linedata, StartQ + 1, EndQ - StartQ - 1) ' Trim the string to only the filename

    If UCase$(Right$(DLLFName, 4)) <> ".DLL" Then DLLFName = DLLFName & ".dll" ' Add a .DLL extension if no extension is present
    If Asc(Right$(DLLFName, 1)) = 34 Then DLLFName = Left$(DLLFName, Len(DLLFName) - 1)  ' Remove excess (junk) characters at end of the string if present

    On Local Error GoTo AlreadyAdded ' An already added error causes it to skip this section

    FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DECDLLS"), tvwChild, "PROJECT_?_DECDLLS_" & UCase$(DLLFName), DLLFName, IsSysDLL(DLLFName)
    DLLRootNodeNum = GetNodeNum("PROJECT_?_DECDLLS_" & UCase$(DLLFName))

    Project.AddData SPF_DecDll, DLLFName

    If FileExists(Environ("windir") & "\System\" & DLLFName) Then
        FileInfo.FindFileInfo Environ("windir") & "\System\" & DLLFName, False
        FrmResults.TreeView.Nodes.Add DLLRootNodeNum, tvwChild, "PROJECT_?_DECDLLS_" & UCase$(DLLFName) & "_PATH", "文件路径: " & Environ("windir") & "\System\" & DLLFName, "Info"
        FrmResults.TreeView.Nodes.Add DLLRootNodeNum, tvwChild, "PROJECT_?_DECDLLS_" & UCase$(DLLFName) & "_SIZE", "文件大小: " & FileInfo.ByteSize, "Info"
        FrmResults.TreeView.Nodes.Add DLLRootNodeNum, tvwChild, "PROJECT_?_DECDLLS_" & UCase$(DLLFName) & "_COMPANYNAME", "公司名称: " & FileInfo.CompanyName, "Info"
        FrmResults.TreeView.Nodes.Add DLLRootNodeNum, tvwChild, "PROJECT_?_DECDLLS_" & UCase$(DLLFName) & "_DESCRIPTION", "描述: " & FileInfo.FileDescription, "Info"
        FrmResults.TreeView.Nodes.Add DLLRootNodeNum, tvwChild, "PROJECT_?_DECDLLS_" & UCase$(DLLFName) & "_VERSION", "版本: " & FileInfo.FileVersion, "Info"
    ElseIf FileExists(Environ("windir") & "\System32\" & DLLFName) Then
        FileInfo.FindFileInfo Environ("windir") & "\System32\" & DLLFName, False
        FrmResults.TreeView.Nodes.Add DLLRootNodeNum, tvwChild, "PROJECT_?_DECDLLS_" & UCase$(DLLFName) & "_PATH", "文件路径: " & Environ("windir") & "\System32\" & DLLFName, "Info"
        FrmResults.TreeView.Nodes.Add DLLRootNodeNum, tvwChild, "PROJECT_?_DECDLLS_" & UCase$(DLLFName) & "_SIZE", "文件大小: " & FileInfo.ByteSize, "Info"
        FrmResults.TreeView.Nodes.Add DLLRootNodeNum, tvwChild, "PROJECT_?_DECDLLS_" & UCase$(DLLFName) & "_COMPANYNAME", "公司名称: " & FileInfo.CompanyName, "Info"
        FrmResults.TreeView.Nodes.Add DLLRootNodeNum, tvwChild, "PROJECT_?_DECDLLS_" & UCase$(DLLFName) & "_DESCRIPTION", "描述: " & FileInfo.FileDescription, "Info"
        FrmResults.TreeView.Nodes.Add DLLRootNodeNum, tvwChild, "PROJECT_?_DECDLLS_" & UCase$(DLLFName) & "_VERSION", "版本: " & FileInfo.FileVersion, "Info"
    End If

AlreadyAdded:
End Sub

Private Sub GetFormStats(linedata As String)  ' Gets the statistics about the form
    Dim FormData As String, JoinLines As Boolean

    Project.ProjectForms = 1 ' Add one (adding function stored in class file to save code)
    CurrSPFLines = 0
    CurrSPFLinesNB = 0

    VBFileNum = FreeFile

    CurrentScanFile "Form" ' Change small picture to a Form image

    If Mid$(linedata, 2, 1) = ":" Then
        If Not FileExists(linedata) Then
            If ShowFNFerrors And Not IsExit Then MsgBoxEx "文件 """ & linedata & """ 未发现!", vbCritical, "扫描错误", , , , , PicError, "Oops!|"
            FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_FORMS"), tvwChild, "PROJECT_?_FORMS_" & linedata, linedata, "Unknown"
            Exit Sub
        Else
            Open linedata For Input As #VBFileNum
        End If
    Else
        If Not FileExists(GetRootDirectory(Project.ProjectPath) & linedata) Then
            If Not FileExists(FilesRootDirectory & GetRootDirectory(Project.ProjectPath) & linedata) Then 'File not found, show an error
                If ShowFNFerrors And Not IsExit Then MsgBoxEx "文件 """ & GetRootDirectory(Project.ProjectPath) & linedata & """ 未发现!", vbCritical, "扫描错误", , , , , PicError, "Oops!|"
                FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_FORMS"), tvwChild, "PROJECT_?_FORMS_" & linedata, linedata, "Unknown"
                Exit Sub
            Else
                Open FilesRootDirectory & GetRootDirectory(Project.ProjectPath) & linedata For Input As #VBFileNum
            End If
        Else
            Open GetRootDirectory(Project.ProjectPath) & linedata For Input As #VBFileNum
        End If
    End If

    Do While Not EOF(VBFileNum)
        Line Input #VBFileNum, FormData ' Get the data from the file
        FormData = Trim$(FormData)

        ' If last line had a line continuation symbol "_", remove the symbol and add the current line to the previous line
        If JoinLines Then
            FormData = Left$(PrjItemPreviousLine, Len(PrjItemPreviousLine) - 1) & FormData
            JoinLines = False
        End If

        If Right$(FormData, 1) = "_" Then ' Line is continued across multiple lines (continuation)
            JoinLines = True ' Tell next line to join to this line
            ProjectItem.PrjItemCodeLines = 1 ' Add to the total lines of code statistic
        Else ' Line is continued across multiple lines (continuation)
            LookAtItemLine FormData ' Examine the line
        End If

        PrjItemPreviousLine = FormData ' Remember the last line, in case of continuation (also used for header check)
    Loop

    Close #VBFileNum

    AddDataToTreeview "FORMS", linedata, "Form", "FRX"

    AddReportText vbNewLine & "============================================================="
    AddReportText "                      VISUAL BASIC FORM"
    AddReportText "============================================================="
    AddReportText "               File Name: " & ExtractFileName(GetRootDirectory(Project.ProjectPath) & linedata)
    AddReportText "                    Name: " & ProjectItem.PrjItemName
    AddReportText vbNewLine & "           Lines (Total): " & ProjectItem.PrjItemCodeLines
    AddReportText "       Lines (No Blanks): " & ProjectItem.PrjItemCodeLinesNoBlanks
    AddReportText "         Lines (Comment): " & ProjectItem.PrjItemCommentLines
    AddReportText "          Lines (Hybrid): " & ProjectItem.PrjItemHybridLines

    AddReportText vbNewLine & "                Controls: " & ProjectItem.PrjItemControls
    AddReportText "               Variables: " & ProjectItem.PrjItemVariables

    AddReportText vbNewLine & "             Subroutines: " & ProjectItem.PrjItemItemSubsCount
    AddReportText "               Functions: " & ProjectItem.PrjItemItemFunctionsCount
    AddReportText "              Properties: " & ProjectItem.PrjItemItemPropertiesCount
    AddReportText "                  Events: " & ProjectItem.PrjItemItemEventsCount
End Sub

Private Sub GetModuleStats(linedata As String)  ' Gets the statistics about the module - see "GetFormStatistics" sub for comments
    Dim ModuleData As String, JoinLines As Boolean, Temp As Long

    Temp = InStr(1, linedata, ";")

    VBFileNum = FreeFile

    If Temp = 0 Then
        MsgBoxEx "当前工程文件错误或丢失数据. 无效行: """ & linedata & """." & vbNewLine & "请使用 Visual Basic 重新打开工程进行修复后再进行扫描.", vbCritical, "扫描错误", , , , , PicError, "Oops!|"
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_MODULES"), tvwChild, "PROJECT_?_MODULES_" & linedata, "(Invalid Data) """ & linedata & """", "Unknown"
        Exit Sub
    End If

    linedata = GetRootDirectory(Left$(linedata, Temp - 1)) & Trim$(Mid$(linedata, Temp + 1))

    Project.ProjectModules = 1
    CurrSPFLines = 0
    CurrSPFLinesNB = 0

    CurrentScanFile "Module"

    If Mid$(linedata, 2, 1) = ":" Then
        If Not FileExists(linedata) Then
            If ShowFNFerrors And Not IsExit Then MsgBoxEx "文件 """ & linedata & """ 未找到!", vbCritical, "扫描错误", , , , , PicError, "Oops!|"
            FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_MODULES"), tvwChild, "PROJECT_?_MODULES_" & linedata, linedata, "Unknown"
            Exit Sub
        Else
            Open linedata For Input As #VBFileNum
        End If
    Else
        If Not FileExists(GetRootDirectory(Project.ProjectPath) & linedata) Then
            If Not FileExists(FilesRootDirectory & GetRootDirectory(Project.ProjectPath) & linedata) Then  'File not found, show an error
                If ShowFNFerrors And Not IsExit Then MsgBoxEx "文件 """ & GetRootDirectory(Project.ProjectPath) & linedata & """ 未找到!", vbCritical, "扫描错误", , , , , PicError, "Oops!|"
                FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_MODULES"), tvwChild, "PROJECT_?_MODULES_" & linedata, linedata, "Unknown"
                Exit Sub
            Else
                Open FilesRootDirectory & GetRootDirectory(Project.ProjectPath) & linedata For Input As #VBFileNum
            End If
        Else
            Open GetRootDirectory(Project.ProjectPath) & linedata For Input As #VBFileNum
        End If
    End If

    Do While Not EOF(VBFileNum)
        Line Input #VBFileNum, ModuleData
        ModuleData = Trim$(ModuleData)

        If JoinLines Then
            ModuleData = Left$(PrjItemPreviousLine, Len(PrjItemPreviousLine) - 1) & ModuleData
            JoinLines = False
        End If

        If Right$(ModuleData, 1) = "_" Then ' Line is continued across multiple lines (continuation)
            JoinLines = True ' Tell next line to join to this line
            ProjectItem.PrjItemCodeLines = 1 ' Add to the total lines of code statistic
        Else ' Line is continued across multiple lines (continuation)
            LookAtItemLine ModuleData ' Examine the line
        End If

        PrjItemPreviousLine = ModuleData ' Remember the last line, in case of continuation (also used for header check)
    Loop

    Close #VBFileNum

    AddDataToTreeview "MODULES", linedata, "Module"

    AddReportText vbNewLine & "============================================================="
    AddReportText "                     VISUAL BASIC MODULE"
    AddReportText "============================================================="
    AddReportText "               File Name: " & ExtractFileName(GetRootDirectory(Project.ProjectPath) & linedata)
    AddReportText "                    Name: " & ProjectItem.PrjItemName
    AddReportText vbNewLine & "           Lines (Total): " & ProjectItem.PrjItemCodeLines
    AddReportText "       Lines (No Blanks): " & ProjectItem.PrjItemCodeLinesNoBlanks
    AddReportText "         Lines (Comment): " & ProjectItem.PrjItemCommentLines
    AddReportText "          Lines (Hybrid): " & ProjectItem.PrjItemHybridLines

    AddReportText vbNewLine & "               Variables: " & ProjectItem.PrjItemVariables

    AddReportText vbNewLine & "             Subroutines: " & ProjectItem.PrjItemItemSubsCount
    AddReportText "               Functions: " & ProjectItem.PrjItemItemFunctionsCount
    AddReportText "              Properties: " & ProjectItem.PrjItemItemPropertiesCount
End Sub

Private Sub GetClassStats(linedata As String)  ' Gets the statistics about the class module - see "GetFormStatistics" sub for comments
    Dim ClassData As String, JoinLines As Boolean, Temp As Long

    Temp = InStr(1, linedata, ";")
    VBFileNum = FreeFile

    If Temp = 0 Then
        MsgBoxEx "当前工程文件错误或丢失数据. 无效行: """ & linedata & """." & vbNewLine & "请使用 Visual Basic 重新打开工程进行修复后再进行扫描.", vbCritical, "扫描错误", , , , , PicError, "Oops!|"
        FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_CLASSES"), tvwChild, "PROJECT_?_CLASSES_" & linedata, "(Invalid Data) """ & linedata & """", "Unknown"
        Exit Sub
    End If

    linedata = GetRootDirectory(Left$(linedata, Temp - 1)) & Trim$(Mid$(linedata, Temp + 1))

    Project.ProjectClasses = 1
    CurrSPFLines = 0
    CurrSPFLinesNB = 0

    CurrentScanFile "Class"

    If Mid$(linedata, 2, 1) = ":" Then
        If Not FileExists(linedata) Then
            If ShowFNFerrors And Not IsExit Then MsgBoxEx "文件 """ & linedata & """ 未发现!", vbCritical, "扫描错误", , , , , PicError, "Oops!|"
            FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_CLASSES"), tvwChild, "PROJECT_?_CLASSES_" & linedata, linedata, "Unknown"
            Exit Sub
        Else
            Open linedata For Input As #VBFileNum
        End If
    Else
        If Not FileExists(GetRootDirectory(Project.ProjectPath) & linedata) Then
            If Not FileExists(FilesRootDirectory & GetRootDirectory(Project.ProjectPath) & linedata) Then  'File not found, show an error
                If ShowFNFerrors And Not IsExit Then MsgBoxEx "文件 """ & GetRootDirectory(Project.ProjectPath) & linedata & """ 未发现!", vbCritical, "扫描错误", , , , , PicError, "Oops!|"
                FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_CLASSES"), tvwChild, "PROJECT_?_CLASSES_" & linedata, linedata, "Unknown"
                Exit Sub
            Else
                Open FilesRootDirectory & GetRootDirectory(Project.ProjectPath) & linedata For Input As #VBFileNum
            End If
        Else
            Open GetRootDirectory(Project.ProjectPath) & linedata For Input As #VBFileNum
        End If
    End If

    Do While Not EOF(VBFileNum)
        Line Input #VBFileNum, ClassData
        ClassData = Trim$(ClassData)

        If JoinLines Then
            ClassData = Left$(PrjItemPreviousLine, Len(PrjItemPreviousLine) - 1) & ClassData
            JoinLines = False
        End If

        If Right$(ClassData, 1) = "_" Then ' Line is continued across multiple lines (continuation)
            JoinLines = True ' Tell next line to join to this line
            ProjectItem.PrjItemCodeLines = 1 ' Add to the total lines of code statistic
        Else ' Line is continued across multiple lines (continuation)
            LookAtItemLine ClassData ' Examine the line
        End If

        PrjItemPreviousLine = ClassData ' Remember the last line, in case of continuation (also used for header check)
    Loop

    Close #VBFileNum

    AddDataToTreeview "CLASSES", linedata, "Class"

    AddReportText vbNewLine & "============================================================="
    AddReportText "                  VISUAL BASIC CLASS MODULE"
    AddReportText "============================================================="
    AddReportText "               File Name: " & ExtractFileName(GetRootDirectory(Project.ProjectPath) & linedata)
    AddReportText "                    Name: " & ProjectItem.PrjItemName
    AddReportText vbNewLine & "           Lines (Total): " & ProjectItem.PrjItemCodeLines
    AddReportText "       Lines (No Blanks): " & ProjectItem.PrjItemCodeLinesNoBlanks
    AddReportText "         Lines (Comment): " & ProjectItem.PrjItemCommentLines
    AddReportText "          Lines (Hybrid): " & ProjectItem.PrjItemHybridLines

    AddReportText vbNewLine & "               Variables: " & ProjectItem.PrjItemVariables

    AddReportText vbNewLine & "             Subroutines: " & ProjectItem.PrjItemItemSubsCount
    AddReportText "               Functions: " & ProjectItem.PrjItemItemFunctionsCount
    AddReportText "              Properties: " & ProjectItem.PrjItemItemPropertiesCount
    AddReportText "                  Events: " & ProjectItem.PrjItemItemEventsCount
End Sub

Private Sub GetUserControlStats(linedata As String)  ' Gets the statistics about the UserControl - see "GetFormStatistics" sub for comments
    Dim UserControlData As String, JoinLines As Boolean

    Project.ProjectUserControls = 1
    CurrSPFLines = 0
    CurrSPFLinesNB = 0

    VBFileNum = FreeFile

    CurrentScanFile "UserControl"

    If Mid$(linedata, 2, 1) = ":" Then
        If Not FileExists(linedata) Then
            If ShowFNFerrors And Not IsExit Then MsgBoxEx "文件 """ & linedata & """ 未发现!", vbCritical, "扫描错误", , , , , PicError, "Oops!|"
            FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_USERCONTROLS"), tvwChild, "PROJECT_?_USERCONTROLS_" & linedata, linedata, "Unknown"
            Exit Sub
        Else
            Open linedata For Input As #VBFileNum
        End If
    Else
        If Not FileExists(GetRootDirectory(Project.ProjectPath) & linedata) Then
            If Not FileExists(FilesRootDirectory & GetRootDirectory(Project.ProjectPath) & linedata) Then 'File not found, show an error
                If ShowFNFerrors And Not IsExit Then MsgBoxEx "文件 """ & GetRootDirectory(Project.ProjectPath) & linedata & """ 未发现!", vbCritical, "扫描错误", , , , , PicError, "Oops!|"
                FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_USERCONTROLS"), tvwChild, "PROJECT_?_USERCONTROLS_" & linedata, linedata, "Unknown"
                Exit Sub
            Else
                Open FilesRootDirectory & GetRootDirectory(Project.ProjectPath) & linedata For Input As #VBFileNum
            End If
        Else
            Open GetRootDirectory(Project.ProjectPath) & linedata For Input As #VBFileNum
        End If
    End If

    Do While Not EOF(VBFileNum)
        Line Input #VBFileNum, UserControlData
        UserControlData = Trim$(UserControlData)

        If JoinLines Then
            UserControlData = Left$(PrjItemPreviousLine, Len(PrjItemPreviousLine) - 1) & UserControlData
            JoinLines = False
        End If

        If Right$(UserControlData, 1) = "_" Then ' Line is continued across multiple lines (continuation)
            JoinLines = True ' Tell next line to join to this line
            ProjectItem.PrjItemCodeLines = 1 ' Add to the total lines of code statistic
        Else ' Line is continued across multiple lines (continuation)
            LookAtItemLine UserControlData ' Examine the line
        End If

        PrjItemPreviousLine = UserControlData ' Remember the last line, in case of continuation (also used for header check)
    Loop

    Close #VBFileNum

    AddDataToTreeview "USERCONTROLS", linedata, "UserControl", "CTX"

    AddReportText vbNewLine & "============================================================="
    AddReportText "                  VISUAL BASIC USER CONTROL"
    AddReportText "============================================================="
    AddReportText "               File Name: " & ExtractFileName(GetRootDirectory(Project.ProjectPath) & linedata)
    AddReportText "                    Name: " & ProjectItem.PrjItemName
    AddReportText vbNewLine & "           Lines (Total): " & ProjectItem.PrjItemCodeLines
    AddReportText "       Lines (No Blanks): " & ProjectItem.PrjItemCodeLinesNoBlanks
    AddReportText "         Lines (Comment): " & ProjectItem.PrjItemCommentLines
    AddReportText "          Lines (Hybrid): " & ProjectItem.PrjItemHybridLines

    AddReportText vbNewLine & "                Controls: " & ProjectItem.PrjItemControls
    AddReportText "               Variables: " & ProjectItem.PrjItemVariables

    AddReportText vbNewLine & "             Subroutines: " & ProjectItem.PrjItemItemSubsCount
    AddReportText "               Functions: " & ProjectItem.PrjItemItemFunctionsCount
    AddReportText "              Properties: " & ProjectItem.PrjItemItemPropertiesCount
    AddReportText "                  Events: " & ProjectItem.PrjItemItemEventsCount
End Sub

Private Sub GetPropertyPageStats(linedata As String)  ' Gets the statistics about the Property Page - see "GetFormStatistics" sub for comments
    Dim PropertyPageData As String, JoinLines As Boolean

    Project.ProjectPropertyPages = 1
    CurrSPFLines = 0
    CurrSPFLinesNB = 0

    VBFileNum = FreeFile

    CurrentScanFile "PropertyPage"

    If Mid$(linedata, 2, 1) = ":" Then
        If Not FileExists(linedata) Then
            If ShowFNFerrors And Not IsExit Then MsgBoxEx "文件 """ & linedata & """ 未发现!", vbCritical, "扫描错误", , , , , PicError, "Oops!|"
            FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_PROPERTYPAGES"), tvwChild, "PROJECT_?_PROPERTYPAGES_" & linedata, linedata, "Unknown"
            Exit Sub
        Else
            Open linedata For Input As #VBFileNum
        End If
    Else
        If Not FileExists(GetRootDirectory(Project.ProjectPath) & linedata) Then
            If Not FileExists(FilesRootDirectory & GetRootDirectory(Project.ProjectPath) & linedata) Then 'File not found, show an error
                If ShowFNFerrors And Not IsExit Then MsgBoxEx "文件 """ & GetRootDirectory(Project.ProjectPath) & linedata & """ 未发现!", vbCritical, "扫描错误", , , , , PicError, "Oops!|"
                FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_PROPERTYPAGES"), tvwChild, "PROJECT_?_PROPERTYPAGES_" & linedata, linedata, "Unknown"
                Exit Sub
            Else
                Open FilesRootDirectory & GetRootDirectory(Project.ProjectPath) & linedata For Input As #VBFileNum
            End If
        Else
            Open GetRootDirectory(Project.ProjectPath) & linedata For Input As #VBFileNum
        End If
    End If

    Do While Not EOF(VBFileNum)
        Line Input #VBFileNum, PropertyPageData
        PropertyPageData = Trim$(PropertyPageData)

        If JoinLines Then
            PropertyPageData = Left$(PrjItemPreviousLine, Len(PrjItemPreviousLine) - 1) & PropertyPageData
            JoinLines = False
        End If

        If Right$(PropertyPageData, 1) = "_" Then ' Line is continued across multiple lines (continuation)
            JoinLines = True ' Tell next line to join to this line
            ProjectItem.PrjItemCodeLines = 1 ' Add to the total lines of code statistic
        Else ' Line is continued across multiple lines (continuation)
            LookAtItemLine PropertyPageData ' Examine the line
        End If

        PrjItemPreviousLine = PropertyPageData ' Remember the last line, in case of continuation (also used for header check)
    Loop

    Close #VBFileNum

    AddDataToTreeview "PROPERTYPAGES", linedata, "PropertyPage"

    AddReportText vbNewLine & "============================================================="
    AddReportText "                  VISUAL BASIC PROPERTY PAGE"
    AddReportText "============================================================="
    AddReportText "               File Name: " & ExtractFileName(GetRootDirectory(Project.ProjectPath) & linedata)
    AddReportText "                    Name: " & ProjectItem.PrjItemName
    AddReportText vbNewLine & "           Lines (Total): " & ProjectItem.PrjItemCodeLines
    AddReportText "       Lines (No Blanks): " & ProjectItem.PrjItemCodeLinesNoBlanks
    AddReportText "         Lines (Comment): " & ProjectItem.PrjItemCommentLines
    AddReportText "          Lines (Hybrid): " & ProjectItem.PrjItemHybridLines

    AddReportText vbNewLine & "                Controls: " & ProjectItem.PrjItemControls
    AddReportText "               Variables: " & ProjectItem.PrjItemVariables

    AddReportText vbNewLine & "             Subroutines: " & ProjectItem.PrjItemItemSubsCount
    AddReportText "               Functions: " & ProjectItem.PrjItemItemFunctionsCount
    AddReportText "              Properties: " & ProjectItem.PrjItemItemPropertiesCount
    AddReportText "                  Events: " & ProjectItem.PrjItemItemEventsCount
End Sub

Private Sub GetDesignerStats(linedata As String)  ' Gets the statistics about the Designer - see "GetFormStatistics" sub for comments
    Dim DesignerData As String, JoinLines As Boolean

    Project.ProjectDesigners = 1
    CurrSPFLines = 0
    CurrSPFLinesNB = 0

    VBFileNum = FreeFile

    CurrentScanFile "Designer"

    If Mid$(linedata, 2, 1) = ":" Then
        If Not FileExists(linedata) Then
            If ShowFNFerrors And Not IsExit Then MsgBoxEx "文件 """ & linedata & """ 未发现!", vbCritical, "扫描错误", , , , , PicError, "Oops!|"
            FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DESIGNERS"), tvwChild, "PROJECT_?_DESIGNERS_" & linedata, linedata, "Unknown"
            Exit Sub
        Else
            Open linedata For Input As #VBFileNum
        End If
    Else
        If Not FileExists(GetRootDirectory(Project.ProjectPath) & linedata) Then
            If Not FileExists(FilesRootDirectory & GetRootDirectory(Project.ProjectPath) & linedata) Then 'File not found, show an error
                If ShowFNFerrors And Not IsExit Then MsgBoxEx "文件 """ & GetRootDirectory(Project.ProjectPath) & linedata & """ 未发现!", vbCritical, "扫描错误", , , , , PicError, "Oops!|"
                FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DESIGNERS"), tvwChild, "PROJECT_?_DESIGNERS_" & linedata, linedata, "Unknown"
                Exit Sub
            Else
                Open FilesRootDirectory & GetRootDirectory(Project.ProjectPath) & linedata For Input As #VBFileNum
            End If
        Else
            Open GetRootDirectory(Project.ProjectPath) & linedata For Input As #VBFileNum
        End If
    End If

    Do While Not EOF(VBFileNum)
        Line Input #VBFileNum, DesignerData
        DesignerData = Trim$(DesignerData)

        If JoinLines Then
            DesignerData = Left$(PrjItemPreviousLine, Len(PrjItemPreviousLine) - 1) & DesignerData
            JoinLines = False
        End If

        If Right$(DesignerData, 1) = "_" Then ' Line is continued across multiple lines (continuation)
            JoinLines = True ' Tell next line to join to this line
            ProjectItem.PrjItemCodeLines = 1 ' Add to the total lines of code statistic
        Else ' Line is continued across multiple lines (continuation)
            LookAtItemLine DesignerData ' Examine the line
        End If

        PrjItemPreviousLine = DesignerData ' Remember the last line, in case of continuation (also used for header check)
    Loop

    Close #VBFileNum

    AddDataToTreeview "DESIGNERS", linedata, "Designer"

    AddReportText vbNewLine & "============================================================="
    AddReportText "                    VISUAL BASIC DESIGNER"
    AddReportText "============================================================="
    AddReportText "               File Name: " & ExtractFileName(GetRootDirectory(Project.ProjectPath) & linedata)
    AddReportText "                    Name: " & ProjectItem.PrjItemName
    AddReportText vbNewLine & "           Lines (Total): " & ProjectItem.PrjItemCodeLines
    AddReportText "       Lines (No Blanks): " & ProjectItem.PrjItemCodeLinesNoBlanks
    AddReportText "         Lines (Comment): " & ProjectItem.PrjItemCommentLines
    AddReportText "          Lines (Hybrid): " & ProjectItem.PrjItemHybridLines

    AddReportText vbNewLine & "                Controls: " & ProjectItem.PrjItemControls
    AddReportText "               Variables: " & ProjectItem.PrjItemVariables

    AddReportText vbNewLine & "             Subroutines: " & ProjectItem.PrjItemItemSubsCount
    AddReportText "               Functions: " & ProjectItem.PrjItemItemFunctionsCount
    AddReportText "              Properties: " & ProjectItem.PrjItemItemPropertiesCount
    AddReportText "                  Events: " & ProjectItem.PrjItemItemEventsCount
End Sub

Private Sub GetUserDocumentStats(linedata As String)  ' Gets the statistics about the User Document - see "GetFormStatistics" sub for comments
    Dim UserDocumentData As String, JoinLines As Boolean

    Project.ProjectUserDocuments = 1
    CurrSPFLines = 0
    CurrSPFLinesNB = 0

    VBFileNum = FreeFile

    CurrentScanFile "UserDocument"

    If Mid$(linedata, 2, 1) = ":" Then
        If Not FileExists(linedata) Then
            If ShowFNFerrors And Not IsExit Then MsgBoxEx "File """ & linedata & """ 未发现!", vbCritical, "扫描错误", , , , , PicError, "Oops!|"
            FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_USERDOCUMENTS"), tvwChild, "PROJECT_?_USERDOCUMENTS_" & linedata, linedata, "Unknown"
            Exit Sub
        Else
            Open linedata For Input As #VBFileNum
        End If
    Else
        If Not FileExists(GetRootDirectory(Project.ProjectPath) & linedata) Then
            If Not FileExists(FilesRootDirectory & GetRootDirectory(Project.ProjectPath) & linedata) Then 'File not found, show an error
                If ShowFNFerrors And Not IsExit Then MsgBoxEx "File """ & GetRootDirectory(Project.ProjectPath) & linedata & """ 未发现!", vbCritical, "扫描错误", , , , , PicError, "Oops!|"
                FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_USERDOCUMENTS"), tvwChild, "PROJECT_?_USERDOCUMENTS_" & linedata, linedata, "Unknown"
                Exit Sub
            Else
                Open FilesRootDirectory & GetRootDirectory(Project.ProjectPath) & linedata For Input As #VBFileNum
            End If
        Else
            Open GetRootDirectory(Project.ProjectPath) & linedata For Input As #VBFileNum
        End If
    End If

    Do While Not EOF(VBFileNum)
        Line Input #VBFileNum, UserDocumentData
        UserDocumentData = Trim$(UserDocumentData)

        If JoinLines Then
            UserDocumentData = Left$(PrjItemPreviousLine, Len(PrjItemPreviousLine) - 1) & UserDocumentData
            JoinLines = False
        End If

        If Right$(UserDocumentData, 1) = "_" Then ' Line is continued across multiple lines (continuation)
            JoinLines = True ' Tell next line to join to this line
            ProjectItem.PrjItemCodeLines = 1 ' Add to the total lines of code statistic
        Else ' Line is continued across multiple lines (continuation)
            LookAtItemLine UserDocumentData ' Examine the line
        End If

        PrjItemPreviousLine = UserDocumentData ' Remember the last line, in case of continuation (also used for header check)
    Loop

    Close #VBFileNum

    AddDataToTreeview "USERDOCUMENTS", linedata, "UserDocument"

    AddReportText vbNewLine & "============================================================="
    AddReportText "                  VISUAL BASIC USER DOCUMENT"
    AddReportText "============================================================="
    AddReportText "               File Name: " & ExtractFileName(GetRootDirectory(Project.ProjectPath) & linedata)
    AddReportText "                    Name: " & ProjectItem.PrjItemName
    AddReportText vbNewLine & "           Lines (Total): " & ProjectItem.PrjItemCodeLines
    AddReportText "       Lines (No Blanks): " & ProjectItem.PrjItemCodeLinesNoBlanks
    AddReportText "         Lines (Comment): " & ProjectItem.PrjItemCommentLines
    AddReportText "          Lines (Hybrid): " & ProjectItem.PrjItemHybridLines

    AddReportText vbNewLine & "                Controls: " & ProjectItem.PrjItemControls
    AddReportText "               Variables: " & ProjectItem.PrjItemVariables

    AddReportText vbNewLine & "             Subroutines: " & ProjectItem.PrjItemItemSubsCount
    AddReportText "               Functions: " & ProjectItem.PrjItemItemFunctionsCount
    AddReportText "              Properties: " & ProjectItem.PrjItemItemPropertiesCount
    AddReportText "                  Events: " & ProjectItem.PrjItemItemEventsCount
End Sub

Private Function GetComponentDescFromReg(ClassID As String) As String  ' Get component description from the registry by it's classID
    Dim RegFDesc As String

    ClassID = Trim$(ClassID)

    RegFDesc = ModRegistry.QueryValue(&H80000000, "TypeLib\" & ClassID, "") 'Find the component's registry folder

    If LenB(RegFDesc) = 0 Then ' Description not found
        GetComponentDescFromReg = "Unknown/Unregistered (ClassID: " & ClassID & ")"
    Else ' Description found, return its value to the calling sub
        GetComponentDescFromReg = RegFDesc
    End If
End Function

Private Function GetComponentNameFromReg(ClassID As String) As String  ' Get the component name from the registry
    Dim RegFName As String

    ClassID = Trim$(ClassID)

    ' Component name should be stored in "HKEY_CLASSES_ROOT\TypeLib\{ClassID}\0\Win32"
    RegFName = ModRegistry.QueryValue(&H80000000, "TypeLib\" & ClassID & "\0\Win32", vbNullString)

    If LenB(RegFName) = 0 Then ' Name not found
        GetComponentNameFromReg = "Unknown/Unregistered (ClassID: " & ClassID & ")"
    Else ' Name found, return its value
        GetComponentNameFromReg = RegFName
    End If
End Function

Private Function CheckIsSourceSafe(ProjectPath As String) As Boolean
    Dim SCCdata As String, IsRightProject As Boolean
    Dim SSafeFileNum As Integer

    ' Source safe data is stored in a file called "MSSCCPRJ.SCC" in the same directory
    ' as the project file - but only if SourceSafe is installed. As a result, don't show
    ' and error if the file isn't there.

    If Not FileExists(GetRootDirectory(ProjectPath) & "MSSCCPRJ.SCC") Then Exit Function

    SSafeFileNum = FreeFile
    Open GetRootDirectory(ProjectPath) & "MSSCCPRJ.SCC" For Input As #SSafeFileNum

    Do While Not EOF(SSafeFileNum)
        Line Input #SSafeFileNum, SCCdata$
        If SCCdata$ = "[" & ExtractFileName(ProjectPath) & "]" Then 'Group files use the same SourceSafe file so find the correct project's info
            IsRightProject = True
        End If

        If IsRightProject And Left$(SCCdata$, 17) = "SCC_Project_Name=" Then  'Data stored in this string
            If Mid$(SCCdata$, 18) <> "this project is not under source code control" Then
                CheckIsSourceSafe = True ' Project is under source safe control
            Else
                CheckIsSourceSafe = False ' Project is not under source safe control
            End If
        End If
    Loop

    Close #SSafeFileNum
End Function

Private Sub LookAtItemLine(linedata As String)  ' Examines a line of data from a VB file
    '                                             This sub is called by all the GetForm/GetModule/etc. subs to save code and increase speed

    Dim SPFName As String, Vars As Long, X As Long, DoNotAddFont As Boolean

    If Not ProjectItem.PrjItemSeenAttributes Then
        If Left$(PrjItemPreviousLine, 10) = "Attribute " Then
            If Left$(linedata, 10) <> "Attribute " And ProjectItem.PrjItemName <> vbNullString Then
                ProjectItem.PrjItemSeenAttributes = True ' Set the flag to true
            End If
        End If

        ' If the file contains control information (Form, User Control, etc.) don't add to statistics until actual code starts
        If ProjectItem.PrjItemInControls Then
            If Left$(linedata, 6) = "Begin " Then
                ProjectItem.PrjItemControls = 1
            ElseIf Left$(linedata, 6) = "Index " Then
                If Val(Mid$(linedata, 18)) Then
                    ProjectItem.PrjItemControlsNoArrays = 1
                End If
            End If

            If Left$(PrjItemPreviousLine, 18) = "BeginProperty Font" Then  ' Control Font
                X = InStr(1, linedata, """") + 1
                SPFName = Mid$(linedata, X, InStrRev(linedata, """") - X)

                For X = 0 To UBound(UsedFonts)
                    If UsedFonts(X) = SPFName Then DoNotAddFont = True
                Next

                If Not DoNotAddFont Then
                    X = UBound(UsedFonts) + 1
                    ReDim Preserve UsedFonts(X) As String
                    UsedFonts(X) = SPFName
                End If
            End If
        End If

        If InStrB(1, linedata, "VB") Then
            If Left$(linedata, 9) = "Begin VB." Then
                ProjectItem.PrjItemInControls = True
            ElseIf Left$(linedata, 20) = "Attribute VB_Name = " Then
                ProjectItem.PrjItemName = Mid$(linedata, 22, Len(linedata) - 22)
                ProjectItem.PrjItemInControls = False ' This is the last line of header info
            End If
        End If

        If Not ProjectItem.PrjItemSeenAttributes Then Exit Sub ' Don't analyse until all the header info is looked at
    End If

    linedata = RemLineNumber(linedata) ' Remove line numbers (if necessary)

    If IsCommentLine(linedata) Then ' Line is a comment line
        ProjectItem.PrjItemCommentLines = 1 ' Add 1 to statistic (adding code is in class to save on space and increase speed)
        Project.ProjectCommentLines = 1 ' Add 1 to total statistic
        CurrSPFLines = CurrSPFLines + 1
        
        Exit Sub ' Don't continue processing line
    Else
        If IsHybridLine(linedata) Then  ' Checks is hybrid code/comment line
            ProjectItem.PrjItemHybridLines = 1 ' Add 1 to statistic (adding code is in class to save on space and increase speed)
            linedata = Mid$(linedata, 1, IsHybridLine(linedata))
        Else
            ProjectItem.PrjItemCodeLines = 1 ' Add 1 to statistic (adding code is in class to save on space and increase speed)
        End If

        If LenB(linedata) = 0 Then
            If LenB(CurrSPFName) = 0 Then ' Not in a SPF, remove the counted blank line from the statistics
                ProjectItem.PrjItemCodeLines = -1
            Else
                CurrSPFLines = CurrSPFLines + 1
            End If

            Exit Sub ' If it's a blank line, don't scan it - it just wastes time
        End If
    End If

    ProjectItem.PrjItemCodeLinesNoBlanks = 1 ' Add 1 to statistic (adding code is in class to save on space and increase speed)

    If InSub Then
        Project.ProjectSubLines = 1
        CurrSPFLinesNB = CurrSPFLinesNB + 1
    ElseIf InFunction Then
        Project.ProjectFuncLines = 1
        CurrSPFLinesNB = CurrSPFLinesNB + 1
    ElseIf InProperty Then
        Project.ProjectPropLines = 1
        CurrSPFLinesNB = CurrSPFLinesNB + 1
    End If

    If CheckIsConstTypeEnum(linedata) Then Exit Sub

    Vars = CheckIsVariable(linedata) ' Check if the line is a variable
    If Vars Then
        ProjectItem.PrjItemVariables = Vars
        Project.ProjectVariables = Vars
        Exit Sub
    End If

    If CheckIsStatement(linedata) Then Exit Sub

    SPFName = CheckIsSub(linedata) ' If the line is a sub, add it to the array for sorting
    If LenB(SPFName) Then
        If Not ShowSPFParams Then SPFName = Left$(SPFName, InStr(1, SPFName, "(") - 1)

        If InStrB(1, SPFName, "Lib") Then
            GetDecDllStats SPFName
            InSub = False
            ProjectItem.AddSPF SPFName & ";N/A [EXT]:N/A", SPF_Sub
            CurrSPFLines = 0
        End If

        CurrSPFName = SPFName
        Exit Sub
    End If

    SPFName = CheckIsFunction(linedata) ' If the line is a function, add it to the list for sorting
    If LenB(SPFName) Then
        If Not ShowSPFParams Then SPFName = Left$(SPFName, InStr(1, SPFName, "(") - 1)

        If InStrB(1, SPFName, "Lib") Then
            GetDecDllStats SPFName
            InFunction = False
            ProjectItem.AddSPF SPFName & ";N/A [EXT]:N/A", SPF_Function
            CurrSPFLines = 0
        End If

        CurrSPFName = SPFName
        Exit Sub
    End If

    SPFName = CheckIsProperty(linedata) ' If the line is a property, add it to the array for sorting
    If LenB(SPFName) Then
        If Not ShowSPFParams Then SPFName = Left$(SPFName, InStr(1, SPFName, "(") - 1)

        CurrSPFName = SPFName
        Exit Sub
    End If

    SPFName = CheckIsEvent(linedata)
    If LenB(SPFName) Then ' If the line is an event, add it to the array for sorting
        If Not ShowSPFParams Then SPFName = Left$(SPFName, InStr(1, SPFName, "(") - 1)

        ProjectItem.AddSPF SPFName, SPF_Event
        IncrementTreeViewPrjEvents
        Exit Sub
    End If

    If CheckForMalicious Then CheckIsMalicious linedata ' If the Potentially Malicious Code checking option is enabled, check the line

    If InStrB(1, linedata, "CreateObject") Then
        X = InStr(1, linedata, "CreateObject") ' CreatObjects are rare, so it's the last SPF thing to check
        If X < 2 Or Mid$(linedata, X - 1, 1) = " " Then ' Correct start for the CreateObject statement
            If Mid$(linedata, X + 12, 1) = "(" Then
                On Local Error Resume Next

                SPFName = Mid$(linedata, X + 13, InStr(X + 13, linedata, ")") - X - 13)
                Project.AddData SPF_CreateObj, SPFName
                FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DECDLLS"), tvwChild, "PROJECT_?_DECDLLS_" & UCase(SPFName), SPFName, "CreateObject"
                FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DECDLLS_" & UCase(SPFName)), tvwChild, "PROJECT_?_DECDLLS_" & UCase(SPFName) & "_INFO", "VB 'CreateObject' Statement", "Info"
                FrmResults.TreeView.Nodes.Add GetNodeNum("PROJECT_?_DECDLLS_" & UCase(SPFName)), tvwChild, "PROJECT_?_DECDLLS_" & UCase(SPFName) & "_INFOLINE", "Code Line: " & linedata, "Info"

                On Local Error GoTo 0
            End If
        End If
    End If
End Sub

Private Function CheckIsSub(linedata As String) As String  ' Returns Sub name if the line is a Sub
    If InStrB(1, linedata, "Sub") Then
        If Left$(linedata, 12) = "Private Sub " Then
            CheckIsSub = Mid$(linedata, 12)
            IncrementTreeViewPrjSubs False
            CurrSPFColour = 1
        ElseIf Left$(linedata, 11) = "Public Sub " Then
            CheckIsSub = Mid$(linedata, 11)
            IncrementTreeViewPrjSubs False
            CurrSPFColour = 2
        ElseIf Left$(linedata, 4) = "Sub " Then
            CheckIsSub = Mid$(linedata, 4)
            IncrementTreeViewPrjSubs False
            CurrSPFColour = 3
        ElseIf Left$(linedata, 20) = "Private Declare Sub " Then
            CheckIsSub = Mid$(linedata, 20)
            IncrementTreeViewPrjSubs True
        ElseIf Left$(linedata, 19) = "Public Declare Sub " Then
            CheckIsSub = Mid$(linedata, 19)
            IncrementTreeViewPrjSubs True
        ElseIf Left$(linedata, 12) = "Declare Sub " Then
            CheckIsSub = Mid$(linedata, 12)
            IncrementTreeViewPrjSubs True
        ElseIf Left$(linedata, 11) = "Friend Sub " Then
            CheckIsSub = Mid$(linedata, 11)
            IncrementTreeViewPrjSubs False
            CurrSPFColour = 4
        ElseIf Left$(linedata, 11) = "Static Sub " Then
            CheckIsSub = Mid$(linedata, 11)
            IncrementTreeViewPrjSubs False
            CurrSPFColour = 5
        ElseIf Left$(linedata, 19) = "Private Static Sub " Then
            CheckIsSub = Mid$(linedata, 19)
            IncrementTreeViewPrjSubs False
            CurrSPFColour = 5
        ElseIf Left$(linedata, 18) = "Public Static Sub " Then
            CheckIsSub = Mid$(linedata, 18)
            IncrementTreeViewPrjSubs False
            CurrSPFColour = 5
        ElseIf Left$(linedata, 18) = "Friend Static Sub " Then
            CheckIsSub = Mid$(linedata, 18)
            IncrementTreeViewPrjSubs False
            CurrSPFColour = 4
        End If

        If LenB(CheckIsSub) Then ' Line is a subroutine header, reset all SPF statistics
            InSub = True
            CurrSPFLines = 0
            CurrSPFLinesNB = 0
            If InStrB(1, CheckIsSub, "'") Then CheckIsSub = Left$(CheckIsSub, InStr(1, CheckIsSub, "'") - 1)
        ElseIf Left$(linedata, 7) = "End Sub" Then
            InSub = False
            ProjectItem.AddSPF CurrSPFName & ";" & (CurrSPFLines + CurrSPFLinesNB - 1) & "[" & CurrSPFColour & "]:" & (CurrSPFLinesNB - 1), SPF_Sub
            CurrSPFName = vbNullString
        End If
    End If
End Function

Private Function CheckIsFunction(linedata As String) As String  ' Returns Function name if the line is a Function
    If InStrB(1, linedata, "Function") Then
        If InStrB(1, linedata, "(") Then
            If InStrB(1, linedata, "Declare ") Then
                If Left$(linedata, 25) = "Private Declare Function " Then
                    CheckIsFunction = Mid$(linedata, 25)
                    IncrementTreeViewPrjFunctions True
                ElseIf Left$(linedata, 24) = "Public Declare Function " Then
                    CheckIsFunction = Mid$(linedata, 24)
                    IncrementTreeViewPrjFunctions True
                ElseIf Left$(linedata, 17) = "Declare Function " Then
                    CheckIsFunction = Mid$(linedata, 17)
                    IncrementTreeViewPrjFunctions True
                End If
            Else
                If Left$(linedata, 8) = "Private " Then
                    If Mid$(linedata, 9, 7) = "Static" Then
                        CheckIsFunction = Mid$(linedata, 25)
                        IncrementTreeViewPrjFunctions False
                        CurrSPFColour = 5
                    Else
                        CheckIsFunction = Mid$(linedata, 17)
                        IncrementTreeViewPrjFunctions False
                        CurrSPFColour = 1
                    End If
                ElseIf Left$(linedata, 7) = "Public " Then
                    If Mid$(linedata, 9, 7) = "Static" Then
                        CheckIsFunction = Mid$(linedata, 24)
                        IncrementTreeViewPrjFunctions False
                        CurrSPFColour = 5
                    Else
                        CheckIsFunction = Mid$(linedata, 16)
                        IncrementTreeViewPrjFunctions False
                        CurrSPFColour = 2
                    End If
                ElseIf Left$(linedata, 9) = "Function " Then
                    CheckIsFunction = Mid$(linedata, 9)
                    IncrementTreeViewPrjFunctions False
                    CurrSPFColour = 3
                ElseIf Left$(linedata, 7) = "Friend " Then
                    CheckIsFunction = Mid$(linedata, 17)
                    IncrementTreeViewPrjFunctions False
                    CurrSPFColour = 4
                ElseIf Left$(linedata, 7) = "Static " Then
                    CheckIsFunction = Mid$(linedata, 17)
                    IncrementTreeViewPrjFunctions False
                    CurrSPFColour = 5
                End If
            End If
        End If

        If LenB(CheckIsFunction) Then ' Line is a function header, reset all SPF statistics
            InFunction = True
            If InStrB(1, CheckIsFunction, "'") Then CheckIsFunction = Left$(CheckIsFunction, InStr(1, CheckIsFunction, "'") - 1)
            CurrSPFName = CheckIsFunction
            CurrSPFLines = 0
            CurrSPFLinesNB = 0
        ElseIf Left$(linedata, 12) = "End Function" Then
            InFunction = False
            ProjectItem.AddSPF CurrSPFName & ";" & (CurrSPFLines + CurrSPFLinesNB - 1) & "[" & CurrSPFColour & "]:" & (CurrSPFLinesNB - 1), SPF_Function
            CurrSPFName = vbNullString
        End If
    End If
End Function

Private Function CheckIsProperty(linedata As String) As String  ' Returns Property name if the line is a Property
    Dim TempStr As String, PLoc As Long

    If InStrB(1, linedata, "Property") Then
        If Left$(linedata, 17) = "Private Property " Then
            TempStr = Mid$(linedata, 18, 3)

            If TempStr = "Let" Then
                CurrSPFColour = 6
            ElseIf TempStr = "Get" Then
                CurrSPFColour = 7
            ElseIf TempStr = "Set" Then
                CurrSPFColour = 8
            End If

            CheckIsProperty = Mid$(linedata, 21)
            IncrementTreeViewPrjProperties
        ElseIf Left$(linedata, 16) = "Public Property " Then
            TempStr = Mid$(linedata, 17, 3)

            If TempStr = "Let" Then
                CurrSPFColour = 6
            ElseIf TempStr = "Get" Then
                CurrSPFColour = 7
            ElseIf TempStr = "Set" Then
                CurrSPFColour = 8
            End If

            CheckIsProperty = Mid$(linedata, 21)
            IncrementTreeViewPrjProperties
        ElseIf Left$(linedata, 9) = "Property " Then
            TempStr = Mid$(linedata, 10, 3)

            If TempStr = "Let" Then
                CurrSPFColour = 6
            ElseIf TempStr = "Get" Then
                CurrSPFColour = 7
            ElseIf TempStr = "Set" Then
                CurrSPFColour = 8
            End If

            CheckIsProperty = Mid$(linedata, 14)
            IncrementTreeViewPrjProperties
        ElseIf Left$(linedata, 16) = "Friend Property " Then
            TempStr = Mid$(linedata, 17, 3)

            If TempStr = "Let" Then
                CurrSPFColour = 6
            ElseIf TempStr = "Get" Then
                CurrSPFColour = 7
            ElseIf TempStr = "Set" Then
                CurrSPFColour = 8
            End If

            CheckIsProperty = Mid$(linedata, 21)
            IncrementTreeViewPrjProperties
        ElseIf Left$(linedata, 7) = "Static " Then
            If InStr(1, linedata, " Property ") Then
                PLoc = InStr(1, linedata, " Property ")
                TempStr = Mid$(linedata, PLoc + 11, 3)

                If TempStr = "Let" Then
                    CurrSPFColour = 6
                ElseIf TempStr = "Get" Then
                    CurrSPFColour = 7
                ElseIf TempStr = "Set" Then
                    CurrSPFColour = 8
                End If

                CheckIsProperty = Mid$(linedata, PLoc + 15)
                IncrementTreeViewPrjProperties
            End If
        End If

        If InStrB(1, CheckIsProperty, "(") = 0 Then CheckIsProperty = ""

        If LenB(CheckIsProperty) Then ' Line is a property header, reset all SPF statistics
            InProperty = True
            CurrSPFLines = 0
            CurrSPFLinesNB = 0
            If InStrB(1, CheckIsProperty, "'") Then CheckIsProperty = Left$(CheckIsProperty, InStr(1, CheckIsProperty, "'") - 1)
            CurrSPFName = CheckIsProperty
        ElseIf Left$(linedata, 12) = "End Property" Then
            InProperty = False
            ProjectItem.AddSPF CurrSPFName & ";" & (CurrSPFLines + CurrSPFLinesNB - 1) & "[" & CurrSPFColour & "]:" & (CurrSPFLinesNB - 1), SPF_Property
            CurrSPFName = vbNullString
        End If
    End If
End Function

Private Function CheckIsEvent(linedata As String)
    If Left$(linedata, 6) = "Event " Then
        CheckIsEvent = Mid$(linedata, 7)
    ElseIf Left$(linedata, 13) = "Public Event " Then
        CheckIsEvent = Mid$(linedata, 14)
    End If
End Function

Private Function CheckIsVariable(linedata As String) As Long  ' Returns TRUE if the line is a Variable
    Dim TempVar As Long

    If Left$(linedata, 4) = "Dim " Then
        CheckIsVariable = 1
    ElseIf Left$(linedata, 7) = "Static " Then
        CheckIsVariable = 1
    ElseIf Left$(linedata, 7) = "Global " Then
        CheckIsVariable = 1
    ElseIf Left$(linedata, 8) = "Private " Then
        CheckIsVariable = 1
    ElseIf Left$(linedata, 7) = "Public " Then
        CheckIsVariable = 1
    ElseIf Left$(linedata, 11) = "WithEvents " Then
        CheckIsVariable = 1
    Else
        Exit Function
    End If

    If InStrB(1, linedata, " Sub ") Then
        CheckIsVariable = 0
        Exit Function
    ElseIf InStrB(1, linedata, " Function ") Then
        CheckIsVariable = 0
        Exit Function
    ElseIf InStrB(1, linedata, " Property ") Then
        CheckIsVariable = 0
        Exit Function
    ElseIf InStrB(1, linedata, " Type ") Then
        CheckIsVariable = 0
        Exit Function
    ElseIf InStrB(1, linedata, " Enum ") Then
        CheckIsVariable = 0
        Exit Function
    ElseIf InStrB(1, linedata, " Event ") Then
        CheckIsVariable = 0
        Exit Function
    ElseIf InStrB(1, linedata, " WithEvents ") Then
        CheckIsVariable = 0
        Exit Function
    ElseIf InStrB(1, linedata, " Const ") Then
        CheckIsVariable = 0
        Exit Function
    End If

    If LineDeclaresArray(linedata) Then
        While InStrB(1, linedata, "(") ' Remove any array information from the variable(s)
            TempVar = InStr(1, linedata, ")")
            linedata = Left$(linedata, InStr(1, linedata, "(") - 1) & Mid$(linedata, TempVar + 1)
        Wend
    End If

    If Left$(linedata, 7) = "Global " Or Left$(linedata, 7) = "Public " Then ' Global variables are added to an array
        If InStrB(1, linedata, "'") Then linedata = Mid$(linedata, 1, InStr(1, linedata, "'") - 1)

        If InStrB(1, linedata, ",") = 0 Then
            GlobalVars_Elements = GlobalVars_Elements + 1
            If UBound(GlobalVars) < GlobalVars_Elements Then
                ReDim Preserve GlobalVars(GlobalVars_Elements + 20) As String
                ReDim Preserve GlobalVarsLoc(GlobalVars_Elements + 20) As String
            End If

            If InStrB(1, linedata, " As ") Then
                GlobalVars(GlobalVars_Elements) = TrimJunk(Mid$(linedata, 8, InStr(1, linedata, " As ") - 8))
                GlobalVarsLoc(GlobalVars_Elements) = ProjectItem.PrjItemName
            Else
                GlobalVars(GlobalVars_Elements) = TrimJunk(Mid$(linedata, 8))
                GlobalVarsLoc(GlobalVars_Elements) = ProjectItem.PrjItemName
            End If
        Else
            Dim VarNames() As String, Var As Long
            Dim UboundVar As Long

            linedata = Mid$(linedata, 8)
            If InStrB(1, linedata, ":") Then linedata = Mid$(linedata, 1, InStr(1, linedata, ":"))

            VarNames = Split(linedata, ",")
            UboundVar = UBound(VarNames) ' Optimisation, for loop is faster comparing against a variable

            For Var = LBound(VarNames) To UboundVar
                GlobalVars_Elements = GlobalVars_Elements + 1
                If UBound(GlobalVars) < GlobalVars_Elements Then
                    ReDim Preserve GlobalVars(GlobalVars_Elements + 20) As String
                    ReDim Preserve GlobalVarsLoc(GlobalVars_Elements + 20) As String
                End If

                If InStrB(1, VarNames(Var), " As ") Then
                    GlobalVars(GlobalVars_Elements) = TrimJunk(Mid$(VarNames(Var), 1, InStr(1, VarNames(Var), " As ")))
                    GlobalVarsLoc(GlobalVars_Elements) = ProjectItem.PrjItemName
                Else
                    GlobalVars(GlobalVars_Elements) = TrimJunk(Mid$(VarNames(Var), 1))
                    GlobalVarsLoc(GlobalVars_Elements) = ProjectItem.PrjItemName
                End If
            Next
        End If
    Else ' Other varibles need less processing
        Dim LineVarLen

        LineVarLen = Len(linedata) ' Optimiation: for loop is faster comparing against a variable

        If InStrB(1, linedata, ",") Then
            For TempVar = 1 To LineVarLen
                If Mid$(linedata, TempVar, 1) = "'" Then Exit For
                If Mid$(linedata, TempVar, 1) = "," Then CheckIsVariable = CheckIsVariable + 1
            Next
        End If
    End If
End Function

Private Function LineDeclaresArray(ByVal strLinedata As String) As Boolean
        If InStr(strLinedata, "(") = 0 Then Exit Function 'If no parenthesis then is not an array so bail now

        strLinedata = Trim$(strLinedata) 'Trim any leading or trailing white space

        'Split the line of text into words.  If the second word (which is element 1 of the array returned
        'by the split function and is the name of our variable) includes a parenthesis then this line is
        'declaring an array
        LineDeclaresArray = CBool(InStr(Split(strLinedata, " ")(1), "("))
End Function

Private Function CheckIsConstTypeEnum(linedata As String) As Boolean
    Dim MidLen As Long

    If Left$(linedata, 7) = "Global " Then
        MidLen = 8
    ElseIf Left$(linedata, 8) = "Private " Then
        MidLen = 9
    ElseIf Left$(linedata, 7) = "Public " Then
        MidLen = 8
    ElseIf Left$(linedata, 7) = "Const " Then
        MidLen = 1
    ElseIf Left$(linedata, 7) = "Type " Then
        MidLen = 1
    ElseIf Left$(linedata, 8) = "Enum " Then
        MidLen = 1
    Else
        Exit Function
    End If

    If Mid$(linedata, MidLen, 5) = "Type " Then
        ProjectItem.PrjItemTypes = 1
        Project.ProjectTypes = 1
        CheckIsConstTypeEnum = True
    ElseIf Mid$(linedata, MidLen, 5) = "Enum " Then
        ProjectItem.PrjItemEnums = 1
        Project.ProjectEnums = 1
        CheckIsConstTypeEnum = True
    ElseIf Mid$(linedata, MidLen, 6) = "Const " Then
        ProjectItem.PrjItemConstants = 1
        Project.ProjectConstants = 1
        CheckIsConstTypeEnum = True
    End If
End Function

Private Function CheckIsStatement(linedata As String) As Boolean
    If Left$(linedata, 3) = "If " Then
        CheckIsStatement = True
        ProjectItem.AddToStatement STIF
    ElseIf Left$(linedata, 4) = "For " Then
        CheckIsStatement = True
        ProjectItem.AddToStatement STFOR
    ElseIf Left$(linedata, 12) = "Select Case " Then
        CheckIsStatement = True
        ProjectItem.AddToStatement STSELECT
    ElseIf Left$(linedata, 3) = "Do " Then
        CheckIsStatement = True
        ProjectItem.AddToStatement STDO
    ElseIf Left$(linedata, 6) = "While " Then
        CheckIsStatement = True
        ProjectItem.AddToStatement STWHILE
    End If
End Function

Public Sub LoadMaliciousKeywords()
    Dim Keyword As String

    Err.Clear ' Clear the error object
    On Local Error GoTo DoneAdd ' Start silent error trapping

    Do
        Keyword = LoadResString(201 + MaliciousKeywordsBuffer_Elements)

        If Keyword <> vbNullString Then
            MaliciousKeywordsBuffer_Elements = MaliciousKeywordsBuffer_Elements + 1 ' Increment the string index variable
            ReDim Preserve MaliciousKeywordsBuffer(MaliciousKeywordsBuffer_Elements) As String
            MaliciousKeywordsBuffer(MaliciousKeywordsBuffer_Elements) = Keyword ' Load the malicious keyword into the buffer
        End If
    Loop

DoneAdd:
    On Local Error GoTo 0 ' Resume no error trapping
End Sub

Private Sub CheckIsMalicious(linedata As String)  ' Checks if a line contains code that could cause trojan or virus like actions
    Dim BLPos As Long, NoAdd As Boolean, LoopVar As Long, I As Long, BCLength As Long

    On Local Error GoTo 0

    ' Since this is not a dedicated PMC scanner, I haven't spent much time making
    ' this - it errs on the side of safety by giving more false positives - I could
    ' extend it but it would be _very_ slow.

    ' Additional note: this only scans for the most common standard commands that would be used
    ' in a malicious way. I have not added checking for API functions (save one or two) because
    ' there just is way too many to check, and I'd never be able to add them all anyway.

    For I = 1 To MaliciousKeywordsBuffer_Elements
        BLPos = InStr(1, linedata, MaliciousKeywordsBuffer(I), vbTextCompare) ' Search for malicious keyword in current line
        BCLength = Len(MaliciousKeywordsBuffer(I))

        If BLPos Then ' Stored malicious keyword found in current line
            If InStrB(1, " ().", Mid$(linedata, IIf(BLPos = 1, 2, BLPos) - 1, 1)) Then
                If (BCLength + BLPos) = Len(linedata) Or InStrB(1, " ().", Mid$(linedata, BLPos + BCLength, 1)) Then
                    NoAdd = False

                    BLPos = UBound(PMaliciousCode) ' Optimisation: For loop is faster comparing against a variable than a property
                    For LoopVar = 0 To BLPos
                        If PMaliciousCode(LoopVar) = linedata Then
                            NoAdd = True
                            Exit For ' Already added, set the variable to TRUE
                        End If
                    Next

                    If Not NoAdd Then  ' Variable FALSE (not added), so add it to the list for sorting
                        BLPos = UBound(PMaliciousCode) + 1
                        ReDim Preserve PMaliciousCode(BLPos) As String
                        PMaliciousCode(BLPos) = linedata

                        Exit Sub ' Found a keyword, don't keep checking the line
                    End If
                End If
            End If
        End If
    Next
End Sub

Private Sub IncrementTreeViewPrjSubs(Declared As Boolean)
    If Declared Then
        NodeNum = GetNodeNum("PROJECT_?_SPF_DECLAREDSUBS") ' Get the node index of the "Total Declared Subs"
    Else
        NodeNum = GetNodeNum("PROJECT_?_SPF_SUBS") ' Get the node index of the "Total Subs"
    End If

    If NodeNum = 0 Then Exit Sub

    ' Conversion to integer for addition is performed automatically:
    FrmResults.TreeView.Nodes(NodeNum).Text = FrmResults.TreeView.Nodes(NodeNum).Text + 1
End Sub

Private Sub IncrementTreeViewPrjFunctions(Declared As Boolean)
    If Declared Then
        NodeNum = GetNodeNum("PROJECT_?_SPF_DECLAREDFUNCTIONS") ' Get the node index of the "Total Declared Functions"
    Else
        NodeNum = GetNodeNum("PROJECT_?_SPF_FUNCTIONS") ' Get the node index of the "Total Functions"
    End If

    If NodeNum = 0 Then Exit Sub

    ' Conversion to integer for addition is performed automatically:
    FrmResults.TreeView.Nodes(NodeNum).Text = FrmResults.TreeView.Nodes(NodeNum).Text + 1
End Sub

Private Sub IncrementTreeViewPrjProperties()
    NodeNum = GetNodeNum("PROJECT_?_SPF_PROPERTIES") ' Get the node index of the "Total Properties"

    If NodeNum = 0 Then Exit Sub

    ' Conversion to integer for addition is performed automatically:
    FrmResults.TreeView.Nodes(NodeNum).Text = FrmResults.TreeView.Nodes(NodeNum).Text + 1
End Sub

Private Sub IncrementTreeViewPrjEvents()
    NodeNum = GetNodeNum("PROJECT_?_SPF_EVENTS") ' Get the node index of the "Total Events"

    If NodeNum = 0 Then Exit Sub

    ' Conversion to integer for addition is performed automatically:
    FrmResults.TreeView.Nodes(NodeNum).Text = FrmResults.TreeView.Nodes(NodeNum).Text + 1
End Sub

Private Sub CurrentScanFile(FileType As String)  ' Changes the little icon on the frmSelProject (if turned on) to
    If ShowCurrItemPic = 0 Then Exit Sub '              indicate what type of file is being scanned

    With FrmSelProject.imgCurrScanObjType
        .Picture = FrmResults.ilstImages.ListImages(FileType).Picture
        .Refresh
    End With
End Sub

Private Sub AddFontsToTreeview() ' Adds found used fonts to the treeview
    Dim LoopVar As Long, TotalItems As Long

    With FrmResults.TreeView.Nodes
        .Add GetNodeNum("PROJECT_?"), tvwChild, "PROJECT_?_FONTS", "使用字体", "Font"
        .Add GetNodeNum("PROJECT_?_FONTS"), tvwChild, "PROJECT_?_FONTS_TOTAL", "总计: " & UBound(UsedFonts), "Info"

        BubbleSortArray UsedFonts ' Sort the array alphabetically before adding it to the treeview

        TotalItems = UBound(UsedFonts)
        For LoopVar = 1 To TotalItems ' This SHOULD be 1 - First array element (index 0) is always blank due to design
            .Add GetNodeNum("PROJECT_?_FONTS"), tvwChild, "PROJECT_?_FONTS_FONT" & LoopVar, UsedFonts(LoopVar), "Font"
        Next
    End With
End Sub

Private Sub AddProjectReportText() ' Add the stats collected about the current project to the report
    Dim StatHeader As String, Temp As String, LoopVar As Long

    StatHeader = "            Startup Item: " & Project.ProjectStartupItem
    StatHeader = StatHeader & vbNewLine & "             Source Safe: " & CheckIsSourceSafe(Project.ProjectPath) & vbNewLine

    StatHeader = StatHeader & vbNewLine & "            Project Type: " & Project.ProjectProjectType
    StatHeader = StatHeader & vbNewLine & "                   Lines: " & Project.ProjectLines
    StatHeader = StatHeader & vbNewLine & "       Lines (No Blanks): " & Project.ProjectLinesNB
    StatHeader = StatHeader & vbNewLine & "         Lines (Comment): " & Project.ProjectCommentLines
    StatHeader = StatHeader & vbNewLine & "      Declared Variables: " & Project.ProjectVariables
    StatHeader = StatHeader & vbNewLine & vbNewLine & "                   Forms: " & Project.ProjectForms

    StatHeader = StatHeader & vbNewLine & "                 Modules: " & Project.ProjectModules
    StatHeader = StatHeader & vbNewLine & "           Class Modules: " & Project.ProjectClasses
    StatHeader = StatHeader & vbNewLine & "           User Controls: " & Project.ProjectUserControls
    StatHeader = StatHeader & vbNewLine & "          User Documents: " & Project.ProjectUserDocuments
    StatHeader = StatHeader & vbNewLine & "          Property Pages: " & Project.ProjectPropertyPages
    StatHeader = StatHeader & vbNewLine & "               Designers: " & Project.ProjectDesigners

    StatHeader = StatHeader & vbNewLine & vbNewLine & "              --- Subs/Functions/Properties ---" & vbNewLine
    StatHeader = StatHeader & vbNewLine & "                    Subs: " & TotalSubs
    StatHeader = StatHeader & vbNewLine & "               Functions: " & TotalFunctions
    StatHeader = StatHeader & vbNewLine & "              Properties: " & TotalProperties
    StatHeader = StatHeader & vbNewLine & "                  Events: " & TotalEvents
    StatHeader = StatHeader & vbNewLine & "      Declared Ext. Subs: " & TotalDecSubs
    StatHeader = StatHeader & vbNewLine & " Declared Ext. Functions: " & TotalDecFunctions

    StatHeader = StatHeader & vbLf & vbNewLine & "                --- Components/References ---" & vbNewLine

    StatHeader = StatHeader & vbNewLine & "-----------------+------------------------------------------+"
    StatHeader = StatHeader & vbNewLine & " FileName:       | Name:                                    |"
    StatHeader = StatHeader & vbNewLine & "-----------------+------------------------------------------+"

    For LoopVar = 1 To Project.ProjectRefComArrayCount Step 2 ' Step 2 to skip SPF description
        Temp = Left$(ExtractFileName(Project.ProjectRefCom(LoopVar)), 16)   ' Gets the filename of the Component/Reference
        If Asc(Right$(Temp, 1)) = 0 Then Temp = Mid$(Temp, 1, Len(Temp) - 1) ' Gets rid of unprintable characters that are sometimes at the end of the strings
        Temp = Temp & Space$(16 - Len(Temp)) ' Adds the spaces if nessesary to preserve the text table formatting
        StatHeader = StatHeader & vbCrLf & Temp & " | " ' Add the text table seperator
        Temp = Left$(Project.ProjectRefCom(LoopVar + 1), 40) ' Get the Component/Reference name
        If Asc(Right$(Temp, 1)) = 0 Then Temp = Mid$(Temp, 1, Len(Temp) - 1) ' Gets rid of unprintable characters that are sometimes at the end of the strings
        StatHeader = StatHeader & Temp & Space$(40 - Len(Temp)) & " |" ' Add the final text table seperator
    Next

    If Project.ProjectRefComArrayCount = 0 Then StatHeader = StatHeader & vbNewLine & "(None)" & Space$(11) & "|" & Space$(42) & "|"

    StatHeader = StatHeader & vbNewLine & "-----------------+------------------------------------------+"

    StatHeader = StatHeader & vbNewLine & vbNewLine & "                    --- Declared DLLs ---" & vbNewLine

    StatHeader = StatHeader & vbNewLine & "------------------------------------------------------------+"
    StatHeader = StatHeader & vbNewLine & " FileName:                                                  |"
    StatHeader = StatHeader & vbNewLine & "------------------------------------------------------------+"

    For LoopVar = 1 To Project.ProjectDecDllsArrayCount
        StatHeader = StatHeader & vbNewLine & Project.ProjectDecDlls(LoopVar) & Space$(59 - Len(Project.ProjectDecDlls(LoopVar))) & " |"
    Next

    If Project.ProjectDecDllsArrayCount = 0 Then StatHeader = StatHeader & vbNewLine & "(None)" & Space$(54) & "|"

    StatHeader = StatHeader & vbNewLine & "------------------------------------------------------------+"

    StatHeader = StatHeader & vbNewLine & vbNewLine & "                  --- Related Documents ---" & vbNewLine
    StatHeader = StatHeader & vbNewLine & "------------------------------------------------------------+"
    StatHeader = StatHeader & vbNewLine & " FileName:                                                  |"
    StatHeader = StatHeader & vbNewLine & "------------------------------------------------------------+"

    If LenB(RelatedDocumentFNames) Then
        StatHeader = StatHeader & vbNewLine & RelatedDocumentFNames
        StatHeader = StatHeader & "------------------------------------------------------------+"
    Else
        StatHeader = StatHeader & vbNewLine & "(None)" & Space$(54) & "|"
        StatHeader = StatHeader & vbNewLine & "------------------------------------------------------------+"
    End If

    FrmReport.rtbReportText.Text = Replace$(FrmReport.rtbReportText.Text, "?PLACE>ProjectName", Project.ProjectName, 1, 1)
    FrmReport.rtbReportText.Text = Replace$(FrmReport.rtbReportText.Text, "?PLACE>ProjectVersion", Project.ProjectVersion, 1, 1)
    FrmReport.rtbReportText.Text = Replace$(FrmReport.rtbReportText.Text, "?PLACE>ProjectStats", StatHeader, 1, 1)
End Sub

Private Function RemLineNumber(linedata As String) As String  ' Remove line numbers from code lines
    Dim LineDataByte As String * 1

    LineDataByte = Left$(linedata, 1)
    If IsNumeric(LineDataByte) Then
        If Int(LineDataByte) = LineDataByte Then
            If InStrB(1, linedata, " ") Then  ' A space after the line number means the line is not blank
                RemLineNumber = Trim$(Mid$(linedata, InStr(1, linedata, " ") + 1)) ' Get line data
            End If
        End If
    Else
        RemLineNumber = linedata
    End If
End Function

Private Function CompareVersions(File1Ver() As String, File2Ver() As String) As String    ' Returns which file is newer from two version numbers
    Dim F1P1 As Integer, F1P2 As Integer, F1P3 As Integer
    Dim F2P1 As Integer, F2P2 As Integer, F2P3 As Integer

    On Local Error Resume Next

    If UBound(File1Ver) > 0 And UBound(File2Ver) > 0 Then
        F1P1 = Int(File1Ver(0)) ' Input is an array,
        F1P2 = Int(File1Ver(1)) ' so split the data
        F1P3 = Int(File1Ver(2)) ' into variables

        F2P1 = Int(File2Ver(0)) ' Input is an array,
        F2P2 = Int(File2Ver(1)) ' so split the data
        F2P3 = Int(File2Ver(2)) ' into variables

        If F1P1 < F2P1 Then CompareVersions = "2N": Exit Function
        If F1P1 > F2P1 Then CompareVersions = "1N": Exit Function

        If F1P2 < F2P2 Then CompareVersions = "2N": Exit Function
        If F1P2 > F2P2 Then CompareVersions = "1N": Exit Function

        If F1P3 < F2P3 Then CompareVersions = "2N": Exit Function
        If F1P3 > F2P3 Then CompareVersions = "1N": Exit Function
    End If

    CompareVersions = "EQ"
End Function

Private Function ZeroIfNull(Data As String) As String ' Returns string zero if the data string is null
    If LenB(Data) Then
        ZeroIfNull = Data
    Else
        ZeroIfNull = "0"
    End If
End Function

Private Function MaskData(Data As String) As String ' Converts string data into masked thousands string data
    If Data <> "0" Then
        MaskData = Format$(Val(Data), "###,###,###")
    Else
        MaskData = Data
    End If
End Function
