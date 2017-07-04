VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmCopyReport 
   BackColor       =   &H00D5E6EA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "复制所需文件的报告"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   4575
   StartUpPosition =   2  '屏幕中心
   Begin DeepLook.ucDeepLookHeader ucDeepLookHeader1 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   661
   End
   Begin MSComctlLib.TreeView tvwNonCopyItemsTV 
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1931
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ilstCopyImages"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ilstCopyImages 
      Left            =   3840
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCopyReport.frx":0000
            Key             =   "Component"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCopyReport.frx":0554
            Key             =   "Project"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCopyReport.frx":095E
            Key             =   "DLL"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCopyReport.frx":0EB2
            Key             =   "SysDLL"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCopyReport.frx":1406
            Key             =   "CreateObject"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCopyReport.frx":1962
            Key             =   "Font"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCopyReport.frx":1CB6
            Key             =   "Error"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCopyReport.frx":2008
            Key             =   "CurrentCopy"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCopyReport.frx":235A
            Key             =   "Done"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCopyReport.frx":26AC
            Key             =   "Skipped"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwItemsTV 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3201
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "ilstCopyImages"
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView tvwManualCopyTV 
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1931
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ilstCopyImages"
      Appearance      =   1
   End
   Begin DeepLook.ucProgressBar pgbPercentBar 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6240
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   12583104
      Scrolling       =   9
      Color2          =   6956042
   End
   Begin DeepLook.ucButtons_H btnStartCopy 
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   6600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "开始复制"
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
      cGradient       =   10395294
      Mode            =   0
      Value           =   0   'False
      Image           =   "FrmCopyReport.frx":2C00
      cBack           =   16777215
   End
   Begin DeepLook.ucButtons_H btnCloseButton 
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   6600
      Width           =   855
      _ExtentX        =   1508
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
      Image           =   "FrmCopyReport.frx":2F54
      cBack           =   16777215
   End
   Begin VB.Label lblManualCopy 
      BackStyle       =   0  'Transparent
      Caption         =   "下列文件可能需要手动复制:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   4455
   End
   Begin VB.Label lblUnessesaryCopyFiles 
      BackStyle       =   0  'Transparent
      Caption         =   "下列文件不需要被复制:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   3855
   End
   Begin VB.Label lblCopyDir 
      BackStyle       =   0  'Transparent
      Caption         =   "(Dir)"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   5760
      Width           =   4215
   End
   Begin VB.Label lblCopyDirLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "复制目录:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label lblFilesBeingCopied 
      BackStyle       =   0  'Transparent
      Caption         =   "下列文件将被复制:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "FrmCopyReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:10:25
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：FrmCopyReport
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:10:25
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V4.12.0
'*************************************************************************

Private Sub btnCloseButton_Click()
    Unload Me
End Sub

Private Sub btnStartCopy_Click()
    Me.Caption = "复制所需文件的报告 - 正在复制中..."
    Me.MousePointer = 13
    btnStartCopy.Enabled = False
    btnCloseButton.Enabled = False
    pgbPercentBar.ShowText = True
    
    On Local Error Resume Next ' Pretection in case directory already created
    MkDir GetRootDirectory(ProjectPath) & "Res\"
    On Local Error GoTo 0
    
    ModXMLReport.MakeXMLCopyReport GetRootDirectory(ProjectPath) & "Res\" & "Copy Report.XML"
    ModFileSearchHandler.CopyDLLOCX
    ModXMLReport.FinishXMLCopyReport GetRootDirectory(ProjectPath) & "Res\" & "Copy Report.XML"
    
    btnCloseButton.Enabled = True
    Me.MousePointer = 1
End Sub

Private Sub Form_Load()
    RemoveTabStops Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then DragForm Me.hwnd ' You can move the forms by holding down the left mouse button
End Sub

Public Sub CheckCheckboxes()
    Dim Count As Long
    
    With tvwItemsTV.Nodes
    For Count = 1 To .Count
        .Item(Count).Checked = True
    Next
    End With
End Sub
