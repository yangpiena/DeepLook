VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmSelProject 
   BackColor       =   &H00D5E6EA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB及.NET工程源代码扫描分析工具 V4.12.0"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "FrmGetProject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmGetProject.frx":23D2
   ScaleHeight     =   2130
   ScaleWidth      =   6120
   StartUpPosition =   2  '屏幕中心
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.ImageCombo cmbProjectPath 
      Height          =   315
      Left            =   825
      TabIndex        =   10
      Top             =   1065
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin DeepLook.ucProgressBar pgbAPB 
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1035
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   49152
      Scrolling       =   9
      Color2          =   6956042
   End
   Begin MSComDlg.CommonDialog cdgDialogs 
      Left            =   5520
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DeepLook.ucButtons_H btnBrowseButton 
      Height          =   300
      Left            =   5160
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1105
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   529
      Caption         =   "..."
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   16777215
      cBhover         =   14737632
      Focus           =   0   'False
      LockHover       =   3
      cGradient       =   14540253
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777215
   End
   Begin DeepLook.ucButtons_H btnGoAnalyse 
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "分析!"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   8421504
      Mode            =   0
      Value           =   0   'False
      Image           =   "FrmGetProject.frx":2714
      Enabled         =   0   'False
      cBack           =   16777215
   End
   Begin DeepLook.ucButtons_H btnOptions 
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "选项"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   8421504
      Mode            =   0
      Value           =   0   'False
      Image           =   "FrmGetProject.frx":2C66
      cBack           =   16777215
   End
   Begin DeepLook.ucButtons_H btnExitDeepLook 
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1680
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   661
      Caption         =   "退出"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   8421504
      Mode            =   0
      Value           =   0   'False
      Image           =   "FrmGetProject.frx":31CA
      cBack           =   16777215
   End
   Begin DeepLook.ucButtons_H btnAbout 
      Height          =   380
      Left            =   4920
      TabIndex        =   9
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "关于"
      CapAlign        =   2
      BackStyle       =   3
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   13323077
      cFHover         =   0
      Focus           =   0   'False
      LockHover       =   2
      cGradient       =   16777215
      Gradient        =   1
      CapStyle        =   2
      Mode            =   0
      Value           =   0   'False
      Image           =   "FrmGetProject.frx":34B0
      cBack           =   16777215
   End
   Begin DeepLook.ucDeepLookHeader hedDeepLookHeader 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   661
   End
   Begin VB.Image imgScanFileIcon 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Picture         =   "FrmGetProject.frx":3A14
      Top             =   720
      Width           =   510
   End
   Begin VB.Label lblScanPhase 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1500
      Width           =   4215
   End
   Begin VB.Image imgCurrScanObjType 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   5160
      Picture         =   "FrmGetProject.frx":42DE
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblScanningName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label lblLocateProjText 
      BackStyle       =   0  'Transparent
      Caption         =   "请选择要扫描分析的 Visual Basic 工程文件:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   5295
   End
End
Attribute VB_Name = "FrmSelProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:10:53
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：FrmSelProject
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:10:53
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V4.12.0
'*************************************************************************

Option Explicit

' -----------------------------------------------------------------------------------------------
Private Declare Function GetTickCount& Lib "kernel32" ()

Private ProjectExt As String
Private Unloading As Boolean
' -----------------------------------------------------------------------------------------------

Private Sub btnAbout_Click()
    FrmAbout.Show 1
End Sub

Private Sub btnBrowseButton_Click()
    cdgDialogs.Filter = "VB6 和 .NET 工程 (*.vbp, *.vbg, *.vbproj)|*.vbp;*.vbg;*.vbproj|Visual Basic 6 工程文件 (*.vbp *.vbg)|*.vbp;*.vbg|Visual Basic 6 文件|*.frm;*.bas;*.cls;*.ctl;*.pag;*.dob;*.dsr|.NET 工程 (*.vbproj)|*.vbproj|工程分析报告 (*.txt)|*.txt|扫描结果文件 (*.dst)|*.dst"
    cdgDialogs.ShowOpen
    If LenB(cdgDialogs.FileName) Then cmbProjectPath.Text = cdgDialogs.FileName
End Sub

Private Sub btnExitDeepLook_Click()
    btnExitDeepLook.Enabled = False
    btnExitDeepLook.Caption = "Wait..."

    IsExit = True
    Unloading = False
    KillProgram False
End Sub

Private Sub Form_Load()
    Load FrmOptions ' Load the options screen, which also loads the options settings

    cmbProjectPath.ImageList = FrmResults.ilstImages ' Set the select project combo listimage to the results main listimage

    Unloading = False
    cmbProjectPath.Visible = True
    pgbAPB.Visible = False
    lblScanningName.Visible = False
    btnGoAnalyse.Enabled = True
    btnBrowseButton.Visible = True
    imgCurrScanObjType.Visible = False
    Me.Caption = "VB及.NET工程源代码扫描分析工具 V4.12.0"
    lblLocateProjText.Caption = "请选择要扫描分析的 Visual Basic 工程文件:"
    lblScanPhase.Caption = ""
    btnOptions.Enabled = True
    Screen.MousePointer = 0
    btnGoAnalyse.Enabled = False

    RemoveTabStops Me
    GetScannedFilesFromRegistry
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then DragForm Me.hwnd ' You can move the forms by holding down the left mouse button
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then ' "X" button or ALT+F4 pressed
        IsExit = True
        KillProgram False
    Else
        KillProgram Unloading
    End If
End Sub

Private Sub btnGoAnalyse_Click()
    GoAnalyse
End Sub

Private Sub btnOptions_Click()
    FrmOptions.Show 1 ' Show the options form (modal 1 so that it must be closed first before continuing)
End Sub

Private Sub cmbProjectPath_Change()
    btnGoAnalyse.Enabled = Len(cmbProjectPath.Text) ' If text entered in the path, enable the scan button

    If Len(cmbProjectPath.Text) > 3 Then
        If Right$(cmbProjectPath.Text, 4) = ".txt" Then
            btnGoAnalyse.Caption = "转换为 XML"
        Else
            btnGoAnalyse.Caption = "分析!"
        End If
    End If
End Sub

Private Sub AddScannedFileToRegistry()
    Dim I As Integer, TotalItems As Long
    Dim HidePath As Boolean, FileName As String

    HidePath = GetSetting("DeepLook", "Options", "HideSelPath", 0)
    
    With cmbProjectPath
        TotalItems = GetSetting("DeepLook", "RecentScanned", "TotalItems", 0)

        For I = 1 To TotalItems
            FileName = GetSetting("DeepLook", "RecentScanned", I, "REGISTRY READ ERROR")
            If FileName = .Text Then Exit Sub
        Next
        
        If .ComboItems.Count <= 75 Then
            SaveSetting "DeepLook", "RecentScanned", "TotalItems", .ComboItems.Count + 1
            SaveSetting "DeepLook", "RecentScanned", .ComboItems.Count + 1, .Text
        End If
        
        BringScannedFileToTop .SelectedItem.Index ' Bring the added file to the top of the list
    End With
End Sub

Public Sub GetScannedFilesFromRegistry()
    Dim TotalItems As Integer, I As Integer
    Dim HidePath As Boolean, FileName As String

    HidePath = GetSetting("DeepLook", "Options", "HideSelPath", 0)

    With cmbProjectPath.ComboItems
        .Clear

        TotalItems = GetSetting("DeepLook", "RecentScanned", "TotalItems", 0)

        On Local Error Resume Next

        For I = 1 To TotalItems
            FileName = GetSetting("DeepLook", "RecentScanned", I, "REGISTRY READ ERROR")
            .Add , FileName, IIf(HidePath, Mid$(FileName, InStrRev(FileName, "\") + 1), FileName), GetPicFromFileTypeExt(Mid$(FileName, InStrRev(FileName, ".") + 1))
        Next

        .Add , , "<< 清除最近的扫描项目/文件记录 >>", "Clean", , 1
    End With
End Sub

Private Sub BringScannedFileToTop(SkipIndex As Integer)
    Dim TotalItems As Integer, I As Integer
    Dim FileName As String

    TotalItems = GetSetting("DeepLook", "RecentScanned", "TotalItems", 0)

    For I = TotalItems To 1 Step -1 ' Run through all the items in the list backwards
        If I <> SkipIndex Then
            FileName = GetSetting("DeepLook", "RecentScanned", I, "REGISTRY READ ERROR")
            SaveSetting "DeepLook", "RecentScanned", I + 1, FileName
        End If
    Next

    SaveSetting "DeepLook", "RecentScanned", 1, cmbProjectPath.Text
End Sub

Private Sub cmbProjectPath_Click()
    Dim RetVal As Integer

    If cmbProjectPath.Text = "<< 清除最近的扫描项目/文件记录 >>" Then
        RetVal = MsgBox("你确定你要清除最近的扫描项目/文件记录?", vbQuestion + vbYesNo + vbDefaultButton2, "提示")

        If RetVal = vbYes Then
            SaveSetting "DeepLook", "RecentScanned", "TotalItems", 0
            GetScannedFilesFromRegistry

            cmbProjectPath.Text = vbNullString
        End If
        
        btnGoAnalyse.Enabled = False
    Else
        If Not cmbProjectPath.SelectedItem Is Nothing Then cmbProjectPath.Text = cmbProjectPath.SelectedItem.Key
    
        cmbProjectPath_Change ' Check to see if the analyse button should be enabled
    End If
End Sub

Public Sub GoAnalyse()
    Dim AddS As String, LinesPerSec As Long, LPSData As String, I As Long
    Dim StartTime As Long, TimeTemp As String

    If cmbProjectPath.Text = "<< 清除最近的扫描项目/文件记录 >>" Then Exit Sub

    If Right$(cmbProjectPath.Text, 4) = ".txt" Then
        ModXMLReport.ConvertTXTtoXML cmbProjectPath.Text
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass ' Set the mouse pointer to busy (hourglass)

    btnOptions.Enabled = False
    btnGoAnalyse.Enabled = False
    btnBrowseButton.Visible = False
    cmbProjectPath.Visible = False

    ReDim UsedFonts(0) As String ' Clears Used Fonts array
    ReDim PMaliciousCode(0) As String ' Clears Potentially Malicious Code array

    TotalLines = 0 ' Reset total scanned lines variable
    ProjectPath = cmbProjectPath.Text ' Get the project/file path
    ProjectExt = Right$(UCase$(cmbProjectPath.Text), 4) ' Get the 3 letter extension of the file to scan
    Me.Caption = "请稍候，系统正在扫描分析中..."

    With FrmResults
        .optStatistics.Enabled = True
        .optUnusedVariables.Enabled = True
        .optCharts.Enabled = True

        On Local Error Resume Next ' Prevent errors from being shown when clearing graphs

        For I = 1 To 3 ' Reset all the Pie Graphs
            .chtFileChart.Column = I
            .chtFileChart.Data = 0
            .chtProjectChart.Column = I
            .chtProjectChart.Data = 0
            .chtSPFChart.Column = IIf(I < 3, I, 1)
            .chtSPFChart.Data = 0
        Next
    End With

    imgCurrScanObjType.Top = 1025
    pgbAPB.Top = 1050
    pgbAPB.Height = 240

    pgbAPB.Visible = True
    lblScanningName.Visible = True
    imgCurrScanObjType.Visible = True

    FrmReport.LockRefresh True ' Lock the report from refreshing, speeds up processing time
    SendMessage FrmResults.TreeView.hwnd, WM_SETREDRAW, 0, 0 ' Same here for the treeview
    FrmReport.rtbReportText.Text = vbNullString ' Clear report

    StartTime = GetTickCount ' Get the scan start time

    AddReportHeader ' Add the header to the report
    ModAnalyseVB6.ClearArrays ' Clear the variable arrays in preperation for the scan
    
    If GetSetting("DeepLook", "Options", "PMCCheck", 1) Then ' If Potentially Malicious Code scanning is enabled
        ModAnalyseVB6.LoadMaliciousKeywords  ' Load the malicious keywords from the RES string table into a buffer array
    End If
    
    pgbAPB.DoubleValOnMetal = False ' Show only the first value on the progressbar
    
    Select Case ProjectExt
        Case ".VBG" ' Visual Basic Group
            lblLocateProjText.Caption = "正在扫描工程组:"
            imgCurrScanObjType.Top = 1000
            pgbAPB.DoubleValOnMetal = True ' Show both values simultaneously
            ModAnalyseVB6.AnalyseGroup cmbProjectPath.Text
            ModAnalyseVB6.PostScanDeleteObjects ' Delete all remaining instances made by the VB6 scanning engine
        Case ".VBP" ' Visual Basic Project
            lblLocateProjText.Caption = "正在扫描工程:"
            ModAnalyseVB6.AnalyseVBProject cmbProjectPath.Text
            ModAnalyseVB6.PostScanDeleteObjects ' Delete all remaining instances made by the VB6 scanning engine
        Case ".FRM", ".BAS", ".CLS", ".CTL", ".PAG", ".DOB", ".DSR"  ' Single Visual Basic
            lblLocateProjText.Caption = "正在扫描文件:"
            ModAnalyseVB6.AnalyseSingleVBItem cmbProjectPath.Text
            ModAnalyseVB6.PostScanDeleteObjects ' Delete all remaining instances made by the VB6 scanning engine
        Case "PROJ" ' Visual Basic.NET Project
            lblLocateProjText.Caption = "正在扫描.NET工程:"
            ModAnalyseDOTNET.AnalyseDotNetProject cmbProjectPath.Text
        Case Else  ' Unknown File Type
            MsgBoxEx "不是Visual Basic 工程文件!", vbCritical, "扫描出错", , , , , PicError, "提示!|"
            Form_Load
            Exit Sub
    End Select

    AddReportFooter ' Add the end information (footer) to the scan report
    FrmReport.rtbReportText.SelStart = 0 ' Put the cursor at the start of the report

    SendMessage FrmResults.TreeView.hwnd, WM_SETREDRAW, 1, 0 ' Unlock the treeview from refreshing
    FrmReport.LockRefresh False ' Unlock the report from refreshing

    Screen.MousePointer = vbNormal ' Set the mousepointer back to normal
    pgbAPB.Value2 = 100 ' Make sure the group progress shows 100% when scan complete

    If FrmResults.TreeView.Nodes.Count < 2 Then ' If only one node in results treeview (only project or group file not found) then don't show results
        Form_Load
        btnExitDeepLook.Enabled = True
        btnGoAnalyse.Enabled = True
        Unload FrmResults
        Exit Sub    ' Don't continue processing the scanned data
    End If

    TimeTemp = Round((GetTickCount - StartTime) / 1000, 2) ' Subtract current time from start time
    If TimeTemp = "1" Then AddS = "" Else AddS = "s" ' Add "s" to the string if scan time is more than 1 second
    If Int(TimeTemp) < 1 Then LinesPerSec = 1 Else LinesPerSec = Int(TimeTemp) ' Lines per second should be a minimum of 1 (can get screwed up if scan time is less than one second)

    LinesPerSec = Round(TotalLines / LinesPerSec, 2) ' Round the lines per second variable to two decimal places
    If LinesPerSec < 1 Then LPSData = "" Else LPSData = " (" & LinesPerSec & " 行/秒)" ' Fix up the lines per second string

    DefaultStatText = "扫描用时 " & TimeTemp & " 秒" & AddS & LPSData & ". 软件当前版本 " & App.Major & "." & App.Minor & "." & App.Revision & "." ' Finish the status bar text string
    FrmResults.sbrStatus.Caption = "STAT>" & DefaultStatText ' Set the status bar text to the created info string

    If PROJECTMODE = VB6 Then ' DeepLook can only copy required files for VB6 files/projects
        FrmResults.btnFileCopy.Enabled = True
        FrmResults.btnSaveUVSList.Enabled = True

        FrmResults.optStatistics.Enabled = True
        FrmResults.optUnusedVariables.Enabled = True
        FrmResults.optCharts.Enabled = True

        If FrmResults.TreeView.Nodes(1).Text = "(Temp Project)" Then
            FrmReport.btnSaveXMLReport.Enabled = False
        Else
            FrmReport.btnSaveXMLReport.Enabled = True
        End If

        With FrmResults.lstVarList
            I = .ListItems.Count

            .ListItems.Add , , ""
            .ListItems.Add , , "总计: " & I ' Add the total number of found unused variables
            .ListItems(.ListItems.Count).Bold = True

            If I Then  ' If unused variables found, show a warning in the treeview
                FrmResults.TreeView.Nodes.Add 1, tvwChild, "未用变量详情", "在扫描工程文件中未使用的变量.", "UnusedVar"
            Else ' No unused variables found
                FrmResults.btnSaveUVSList.Enabled = False ' Disable the save unused variable list button
                FrmResults.optUnusedVariables.Enabled = False ' Disable the unused variable list view button
                FrmResults.TreeView.Nodes.Add 1, tvwChild, "未用变量详情", "在扫描工程文件中未使用的变量.", "NoUnusedVar"
            End If
        End With
    Else ' .NET projects cannot copy dependancies or show pie graphs/unused variables
        With FrmResults
            .optStatistics.Enabled = False
            .optUnusedVariables.Enabled = False
            .optCharts.Enabled = False

            .btnSaveUVSList.Enabled = False
            .btnFileCopy.Enabled = False
        End With

        FrmReport.btnSaveXMLReport.Enabled = False
    End If

    With FrmResults
        If .TreeView.Nodes.Count > 16000 Then ' Show large number of nodes warning in the treeview
            .TreeView.Nodes.Add 1, tvwChild, "警告节点", "Warning: There are a large amount (" & Format$(FrmResults.TreeView.Nodes.Count, "###,###,###") & ") of nodes in the treeview.", "警告"
            .TreeView.Nodes.Add 1, tvwChild, "警告节点2", "This large number may cause DeepLook or your computer to run slowly or stop responding."
        End If

        If GetSetting("DeepLook", "Options", "AllowOnlyOneBranch", 0) Then
            .TreeView.SingleSel = True
            .btnExpand.Enabled = False
            .btnExpand.Height = 365 ' Quirk of the LaVolpe button, disabled looks taller, must shrink the size to compensate
            
            .mnuShowall.Enabled = False '  Can't show all items if only one branch allowed to be open at one time,
            .mnuShowall2.Enabled = False ' so disable the appropriate menu items
        Else
            .TreeView.SingleSel = False
            .btnExpand.Enabled = True
            .btnExpand.Height = 375
        End If

        If Not IsExit Then ' Not trying to quit (exit DeepLook button pressed)
            Unloading = True ' Set the unlaoding flag to true (if form terminated and flag false, program is closed)
            .Show ' Show the results form
        End If
    End With

    AddScannedFileToRegistry ' Add the scanned file to the scanned file list if nessesary
    If GetSetting("DeepLook", "Options", "BringScannedFileToTop", 1) Then
        BringScannedFileToTop cmbProjectPath.ComboItems.Count ' Now bring the added file to the top of the list
    End If
    
    Unload Me ' Unload the select project form
End Sub

Private Function GetPicFromFileTypeExt(Ext As String)
    Select Case UCase$(Ext)
        Case "VBP"
            GetPicFromFileTypeExt = "Project"
        Case "VBG"
            GetPicFromFileTypeExt = "Group"
        Case "VBPROJ"
            GetPicFromFileTypeExt = "NETproject"
        Case "VB"
            GetPicFromFileTypeExt = "NETvb"
        Case "FRM"
            GetPicFromFileTypeExt = "Form"
        Case "BAS"
            GetPicFromFileTypeExt = "Module"
        Case "CLS"
            GetPicFromFileTypeExt = "Class"
        Case "CTL"
            GetPicFromFileTypeExt = "UserControl"
        Case "DSR"
            GetPicFromFileTypeExt = "Designer"
        Case "DOB"
            GetPicFromFileTypeExt = "UserDocument"
    End Select
End Function
