VERSION 5.00
Begin VB.Form FrmAbout 
   BackColor       =   &H00D5E6EA&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About DeepLook"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5040
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   5040
   StartUpPosition =   2  '屏幕中心
   Begin DeepLook.ucButtons_H btnCloseButton 
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   6360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Close"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   8421504
      Mode            =   0
      Value           =   0   'False
      Image           =   "FrmAbout.frx":06EA
      cBack           =   16777215
   End
   Begin DeepLook.ucDeepLookHeader hedDeepLookHeader 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5220
      _ExtentX        =   9049
      _ExtentY        =   661
   End
   Begin VB.Frame fmeAbout 
      BackColor       =   &H00D5E6EA&
      Caption         =   "About DeepLook"
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4815
      Begin DeepLook.ucThreeDLine linSep2 
         Height          =   45
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   79
      End
      Begin DeepLook.ucThreeDLine linSep4 
         Height          =   45
         Left            =   120
         TabIndex        =   14
         Top             =   3480
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   79
      End
      Begin DeepLook.ucThreeDLine linSep5 
         Height          =   45
         Left            =   120
         TabIndex        =   15
         Top             =   3840
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   79
      End
      Begin DeepLook.ucThreeDLine linSep6 
         Height          =   45
         Left            =   120
         TabIndex        =   16
         Top             =   4320
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   79
      End
      Begin DeepLook.ucThreeDLine linSep7 
         Height          =   45
         Left            =   120
         TabIndex        =   17
         Top             =   5040
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   79
      End
      Begin DeepLook.ucThreeDLine linSep1 
         Height          =   45
         Left            =   1440
         TabIndex        =   12
         Top             =   1080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   79
      End
      Begin DeepLook.ucThreeDLine linSep8 
         Height          =   45
         Left            =   120
         TabIndex        =   20
         Top             =   5400
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   79
      End
      Begin VB.Label lblMinRes 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DeepLook requires a minimum screen resolution of 1024x768."
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
         Left            =   120
         TabIndex        =   19
         Top             =   5160
         Width           =   4575
      End
      Begin VB.Image imgAuthor 
         Height          =   1380
         Left            =   120
         Picture         =   "FrmAbout.frx":09D0
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label lblControlCredit2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmAbout.frx":3EA8
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
         Left            =   120
         TabIndex        =   18
         Top             =   2880
         Width           =   4695
      End
      Begin VB.Label lblEmail 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Email me at dean_camera@hotmail.com."
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
         Left            =   120
         TabIndex        =   10
         Top             =   5520
         Width           =   4575
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmAbout.frx":3F6C
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
         Left            =   195
         TabIndex        =   9
         Top             =   4440
         Width           =   4455
      End
      Begin VB.Label lblSpecialThanks 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Special thanks to the members of PlanetSourceCode.com who's suggestions have helped in the development of DeepLook."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3930
         Width           =   4575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DeepLook VB Project Scanner"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00D5E6EA&
         Caption         =   "By Dean Camera, 2003-2005"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "This is DeepLook version #.#.# (built #)."
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
         Left            =   120
         TabIndex        =   3
         Top             =   3600
         Width           =   4575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:  Nationality:   Age:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00D5E6EA&
         BackStyle       =   0  'Transparent
         Caption         =   "Dean Camera   Australian      16"
         Height          =   615
         Left            =   2520
         TabIndex        =   1
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblControlCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmAbout.frx":403E
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   60
         TabIndex        =   7
         Top             =   2040
         Width           =   4695
      End
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:08:59
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：FrmAbout
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:08:59
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V4.12.0
'*************************************************************************
Option Explicit

Dim RunningInIDE As Boolean

Private Sub btnCloseButton_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    Dim EXEInfo As ClsFileProp, EXEPath As String, BuildData As String
    
    On Local Error Resume Next
    
    Set EXEInfo = New ClsFileProp
    RemoveTabStops Me

    If Right$(App.Path, 1) = "\" Then
        EXEPath = App.Path & App.EXEName
    Else
        EXEPath = App.Path & "\" & App.EXEName
    End If

    Debug.Assert IDECheck ' All Debug statements are ignored/stripped upon compile, thus the sub is never
    '                       run when compiled and so the RunningInIde variable stays false

    If RunningInIDE Then ' Running inside the IDE (not compiled)
        BuildData = " (Running in IDE)"
    Else
        EXEInfo.FindFileInfo EXEPath & ".exe", False ' Try to get the EXE creation stats
        BuildData = " (Built " & Mid$(EXEInfo.CreationTime, 1, InStrRev(EXEInfo.CreationTime, " ") - 1) & ")"
    End If
    
    Set EXEInfo = Nothing

    lblVersion.Caption = "This is DeepLook version " & App.Major & "." & App.Minor & "." & App.Revision & BuildData & "."
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then DragForm Me.hwnd ' You can move the forms by holding down the left mouse button
End Sub

Private Function IDECheck() ' Dummy function, sets the RunningInIDE variable when executed
    IDECheck = 1 ' Must set the return to non-zero, other wise the Debug.Assert statement halts the program
    RunningInIDE = True ' Running inside the IDE, so set the variable accordingly
End Function
