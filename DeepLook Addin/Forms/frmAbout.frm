VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00D5E6EA&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "关于插件"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fmeAbout 
      BackColor       =   &H00D5E6EA&
      Caption         =   "About DeepLook Addin"
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin DeepLookAddin.ucThreeDLine linSep6 
         Height          =   45
         Left            =   120
         TabIndex        =   15
         Top             =   3720
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   79
         LineColour      =   0
      End
      Begin DeepLookAddin.ucThreeDLine linSep5 
         Height          =   45
         Left            =   120
         TabIndex        =   14
         Top             =   3360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   79
         LineColour      =   0
      End
      Begin DeepLookAddin.ucThreeDLine linSep4 
         Height          =   45
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   79
         LineColour      =   0
      End
      Begin DeepLookAddin.ucThreeDLine linSep3 
         Height          =   45
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   79
         LineColour      =   0
      End
      Begin DeepLookAddin.ucThreeDLine linSep2 
         Height          =   45
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   79
         LineColour      =   0
      End
      Begin DeepLookAddin.ucThreeDLine linSep1 
         Height          =   45
         Left            =   1440
         TabIndex        =   5
         Top             =   840
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   79
         LineColour      =   0
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "This is DeepLook Addin version #.#.#."
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
         TabIndex        =   11
         Top             =   1920
         Width           =   4575
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
         TabIndex        =   10
         Top             =   2250
         Width           =   4575
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":0000
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
         Top             =   2760
         Width           =   4455
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
         TabIndex        =   8
         Top             =   3840
         Width           =   4575
      End
      Begin VB.Label lblMinReq 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DeepLook Addin requires a compiled binary of DeepLook to function."
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
         TabIndex        =   7
         Top             =   3480
         Width           =   4575
      End
      Begin VB.Label Label8 
         BackColor       =   &H00D5E6EA&
         BackStyle       =   0  'Transparent
         Caption         =   "Dean Camera   Australian      16"
         Height          =   615
         Left            =   2520
         TabIndex        =   4
         Top             =   960
         Width           =   1095
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
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00D5E6EA&
         Caption         =   "By Dean Camera, 2005"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   560
         Width           =   3255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DeepLook VB Project Scanner Addin"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1370
         TabIndex        =   1
         Top             =   240
         Width           =   3495
      End
      Begin VB.Image imgAuthor 
         Height          =   1380
         Left            =   120
         Picture         =   "frmAbout.frx":00D2
         Top             =   240
         Width           =   1200
      End
   End
   Begin DeepLookAddin.ucButtons_H btnClose 
      Height          =   375
      Left            =   1920
      TabIndex        =   16
      Top             =   4320
      Width           =   1095
      _ExtentX        =   1931
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
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmAbout.frx":35AA
      cBack           =   16777215
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:15:10
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：frmAbout
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:15:10
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V1.0.0
'*************************************************************************
Option Explicit

Private Sub btnClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Hide
End Sub
