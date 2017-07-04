VERSION 5.00
Begin VB.Form SearchReg 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "文件注册"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4455
   Icon            =   "FrmSearchReg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton BtnExit 
      Caption         =   "退出"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.FileListBox RegFiles 
      Height          =   630
      Left            =   120
      Pattern         =   "*.ocx;*.dll;*.tlb"
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Info 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "SearchReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:16:42
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：插件注册
'**模 块 名：SearchReg
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:16:42
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V1.0.0
'*************************************************************************
Dim FileSearch As New ClsSearch
Attribute FileSearch.VB_VarHelpID = -1

Private Sub BtnExit_Click()
    End
End Sub

Private Sub Form_Load()
    RegFiles.Path = App.Path
    Me.Show
    DoEvents
    Info.Text = Info.Text & "Searching for RegSvr32.exe...."
    DoEvents
    FileSearch.SearchFiles Environ("windir"), "RegSvr32.exe", True

    If FileSearch.Files.Count <> 0 Then
        Info.Text = Info.Text & "Found."

        For i = 0 To RegFiles.ListCount - 1
            Info.Text = Info.Text & vbNewLine & "Registering " & RegFiles.List(i) & "..."
            Shell FileSearch.Files.Item(1).FilePath & "RegSvr32.exe /s " & RegFiles.List(i)
            Info.Text = Info.Text & "Done."
            DoEvents
        Next

        Info.Text = Info.Text & vbNewLine & "All files registered."
    Else
        Info.Text = Info.Text & "Not Found, file reg failed."
    End If
End Sub
