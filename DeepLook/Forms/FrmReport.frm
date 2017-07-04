VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form FrmReport 
   BackColor       =   &H00D5E6EA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "工程报告"
   ClientHeight    =   7260
   ClientLeft      =   4200
   ClientTop       =   2550
   ClientWidth     =   7605
   Icon            =   "FrmReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   7605
   StartUpPosition =   2  '屏幕中心
   Begin DeepLook.ucButtons_H btnTextReport 
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Text"
      CapAlign        =   2
      BackStyle       =   1
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   12632256
      cGradient       =   12632256
      Gradient        =   3
      Mode            =   2
      Value           =   0   'False
      Image           =   "FrmReport.frx":058A
      cBack           =   16777215
   End
   Begin DeepLook.ucButtons_H btnXMLReport 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "XML"
      CapAlign        =   2
      BackStyle       =   1
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   12632256
      cGradient       =   12632256
      Gradient        =   3
      Mode            =   2
      Value           =   -1  'True
      Image           =   "FrmReport.frx":069C
      cBack           =   16777215
   End
   Begin DeepLook.ucDeepLookHeader ucDeepLookHeadder1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7830
      _ExtentX        =   13573
      _ExtentY        =   661
   End
   Begin MSComDlg.CommonDialog cdgCommonDialog 
      Left            =   4920
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DeepLook.ucButtons_H btnSaveReport 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   6840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "另存为 TXT 文件"
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
      Image           =   "FrmReport.frx":0C36
      cBack           =   16777215
   End
   Begin DeepLook.ucButtons_H btnCloseButton 
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   6840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "关闭"
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
      Image           =   "FrmReport.frx":118A
      cBack           =   16777215
   End
   Begin DeepLook.ucButtons_H btnSaveXMLReport 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   6840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "另存为 XML 文件"
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
      Image           =   "FrmReport.frx":1470
      cBack           =   16777215
   End
   Begin SHDocVwCtl.WebBrowser wbrBrowser 
      Height          =   5895
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   7335
      ExtentX         =   12938
      ExtentY         =   10398
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin RichTextLib.RichTextBox rtbReportText 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   10398
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"FrmReport.frx":19C4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:10:41
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：工程报告
'**模 块 名：FrmReport
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:10:41
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V4.12.0
'*************************************************************************

Option Explicit

Public Sub LockRefresh(YesNo As Boolean)
    SendMessage rtbReportText.hwnd, WM_SETREDRAW, (Not YesNo), 0
End Sub

Private Sub btnCloseButton_Click()
    Me.Visible = False
End Sub

Private Sub btnSaveReport_Click()
    Dim RetVal As Integer

    With cdgCommonDialog
        .Filter = "DeepLook Report File (*.txt)|*.txt"
        .FileName = vbNullString
        .ShowSave

        If .FileName = vbNullString Then Exit Sub

        If FileExists(.FileName) Then
            RetVal = MsgBoxEx("文件 """ & .FileName & """ 已经存在. 请问是否覆盖?", vbExclamation Or vbYesNo Or vbDefaultButton2, "提示", , , , , PicReport)
            If RetVal = vbNo Then Exit Sub
        End If

        rtbReportText.SaveFile .FileName, rtfText
    End With

    btnTextReport.Value = True
    rtbReportText.Visible = True
    wbrBrowser.Visible = False
End Sub

Private Sub btnSaveXMLReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim RetVal As Integer, I As Integer

    With cdgCommonDialog
        .Filter = " XML 报告文件 (*.xml)|*.xml"
        .FileName = vbNullString
        .ShowSave

        If .FileName = vbNullString Then Exit Sub

        If FileExists(.FileName) Then
            RetVal = MsgBoxEx("文件 """ & .FileName & """ 已经存在. 是否覆盖?", vbExclamation Or vbYesNo Or vbDefaultButton2, "提示", , , , , PicReport)
            If RetVal = vbNo Then Exit Sub
        End If

        rtbReportText.SaveFile .FileName & ".txt", rtfText
        ModXMLReport.MakeXMLReport .FileName, .FileName & ".txt" ' Generate the XML report

        RetVal = FreeFile
        Open GetRootDirectory(.FileName) & "DeepLook.xsl" For Output As #RetVal ' Open a blank XML stylesheet

        Dim StrData As String
        StrData = StrConv(LoadResData(1, "XMLTEMPLATE"), vbUnicode)
        StrData = IIf(Right(StrData, 2) <> "t>", Left$(StrData, Len(StrData) - 2), StrData)

        Print #RetVal, StrData ' Write the XML stylesheet into the created file

        Close #RetVal ' Close the XML Stylesheet file

        Kill .FileName & ".txt"  ' Delete the tempoary text report file

        btnXMLReport.Value = True
        rtbReportText.Visible = False
        wbrBrowser.Visible = True
    End With
End Sub

Private Sub btnTextReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    rtbReportText.Visible = True
    wbrBrowser.Visible = False
End Sub

Private Sub btnXMLReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    rtbReportText.Visible = False
    wbrBrowser.Visible = True
End Sub

Private Sub Form_Load()
    RemoveTabStops Me

    wbrBrowser.Navigate "about:Loading..."
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then DragForm Me.hwnd ' You can move the forms by holding down the left mouse button
End Sub

Public Sub CreateTempXMLReport()
    Dim StrData As String
    Dim RetVal As Integer
    Dim TempPath As String

    If PROJECTMODE = VB6 Then
        TempPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")

        rtbReportText.SaveFile TempPath & "DeepLookXMLTemp.xml" & ".txt", rtfText
        ModXMLReport.MakeXMLReport TempPath & "DeepLookXMLTemp.xml", TempPath & "DeepLookXMLTemp.xml" & ".txt" ' Generate the XML report

        RetVal = FreeFile
        Open TempPath & "DeepLook.xsl" For Output As #RetVal ' Open a blank XML stylesheet

        StrData = StrConv(LoadResData(1, "XMLTEMPLATE"), vbUnicode)
        StrData = IIf(Right(StrData, 2) <> "t>", Left$(StrData, Len(StrData) - 2), StrData)

        Print #RetVal, StrData ' Write the XML stylesheet into the created file
        Close #RetVal ' Close the XML Stylesheet file

        Kill TempPath & "DeepLookXMLTemp.xml" & ".txt"  ' Delete the tempoary text report file

        wbrBrowser.Navigate TempPath & "DeepLookXMLTemp.xml"
    Else
        wbrBrowser.Navigate "about:XML Report not avaliable for .NET."
        
        btnTextReport.Value = True ' Make the text report button selected
        btnTextReport_MouseDown 1, 1, 1, 1 ' Show the text report
    End If
End Sub
