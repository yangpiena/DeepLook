VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form FrmResults 
   BackColor       =   &H00D5E6EA&
   Caption         =   "工程扫描分析结果"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13215
   Icon            =   "FrmResults.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   13215
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ImageList ilstImages 
      Left            =   840
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12583104
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   38
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":058A
            Key             =   "Project"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":0996
            Key             =   "Group"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":0D5E
            Key             =   "DLL"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":12B2
            Key             =   "SysDLL"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":1806
            Key             =   "Component"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":1D5A
            Key             =   "REFCOM"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":22AE
            Key             =   "RelatedDocuments"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":2802
            Key             =   "Form"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":2B96
            Key             =   "Module"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":2FAA
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":33BE
            Key             =   "UserControl"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":380E
            Key             =   "UserDocument"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":3BA2
            Key             =   "PropertyPage"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":3FB6
            Key             =   "Method"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":43CE
            Key             =   "Property"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":47E2
            Key             =   "BadCode"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":48F6
            Key             =   "Clean"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":4C8A
            Key             =   "Unknown"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":4FDE
            Key             =   "SPF"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":537A
            Key             =   "Resource"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":5920
            Key             =   "RelDoc"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":5A8A
            Key             =   "NETproject"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":5DDE
            Key             =   "NETvb"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":6132
            Key             =   "Info"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":7186
            Key             =   "LOGFile"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":755A
            Key             =   "IComponent"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":7AAE
            Key             =   "App"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":8006
            Key             =   "Designer"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":841E
            Key             =   "CodeLoop"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":8532
            Key             =   "Total"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":8886
            Key             =   "Event"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":89DA
            Key             =   "CreateObject"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":8F36
            Key             =   "Warning"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":928A
            Key             =   "Font"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":95DE
            Key             =   "UnusedVar"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":9930
            Key             =   "NoUnusedVar"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":9C82
            Key             =   "ProjStats"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":9FD4
            Key             =   "ProjFileInfo"
         EndProperty
      EndProperty
   End
   Begin DeepLook.ucButtons_H optCharts 
      Height          =   315
      Left            =   3720
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   405
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "图表"
      CapAlign        =   2
      BackStyle       =   1
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
      cBhover         =   12632256
      Focus           =   0   'False
      cGradient       =   12632256
      Gradient        =   3
      Mode            =   2
      Value           =   0   'False
      Image           =   "FrmResults.frx":A326
      cBack           =   16777215
   End
   Begin MSComDlg.CommonDialog cdgDialogs 
      Left            =   240
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DeepLook.ucButtons_H sbrStatus 
      Height          =   315
      Left            =   0
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   6840
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   556
      Caption         =   "STAT>Status Bar"
      CapAlign        =   2
      BackStyle       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cBhover         =   14737632
      cGradient       =   14737632
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   16777215
   End
   Begin DeepLook.ucButtons_H btnAbout 
      Height          =   375
      Left            =   12000
      TabIndex        =   44
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
      Image           =   "FrmResults.frx":A5E8
      cBack           =   16777215
   End
   Begin MSComctlLib.ImageList ilstUnusedVarImages 
      Left            =   1560
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":AB4C
            Key             =   "Name"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":AC60
            Key             =   "File"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":ADCC
            Key             =   "Project"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmResults.frx":B1D8
            Key             =   "GlobalLocal"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pbxControlPanel 
      BackColor       =   &H00D5E6EA&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   360
      ScaleHeight     =   375
      ScaleWidth      =   12495
      TabIndex        =   2
      Top             =   6360
      Width           =   12495
      Begin DeepLook.ucButtons_H btnCollapseAll 
         Height          =   375
         Left            =   815
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "项目"
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
         cBhover         =   12632256
         Focus           =   0   'False
         cGradient       =   12632256
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "FrmResults.frx":B2EC
         cBack           =   16777215
      End
      Begin DeepLook.ucThreeDLine linSep2 
         Height          =   375
         Left            =   9840
         TabIndex        =   29
         Top             =   0
         Width           =   45
         _ExtentX        =   79
         _ExtentY        =   661
      End
      Begin DeepLook.ucThreeDLine linSep1 
         Height          =   375
         Left            =   2130
         TabIndex        =   33
         Top             =   0
         Width           =   45
         _ExtentX        =   79
         _ExtentY        =   661
      End
      Begin DeepLook.ucButtons_H btnFileCopy 
         Height          =   375
         Left            =   2280
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   0
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         Caption         =   "复制支撑所需文件"
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
         Image           =   "FrmResults.frx":B83E
         cBack           =   16777215
      End
      Begin DeepLook.ucButtons_H btnGenerateReport 
         Height          =   375
         Left            =   4800
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   0
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         Caption         =   "显示工程报告"
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
         Image           =   "FrmResults.frx":BB92
         cBack           =   16777215
      End
      Begin DeepLook.ucButtons_H btnScanAnother 
         Height          =   375
         Left            =   9960
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   0
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "扫描其他工程"
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
         Image           =   "FrmResults.frx":C116
         cBack           =   16777215
      End
      Begin DeepLook.ucButtons_H btnExitDeepLook 
         Height          =   375
         Left            =   11640
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
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
         Image           =   "FrmResults.frx":C46A
         cBack           =   16777215
      End
      Begin DeepLook.ucButtons_H btnExpand 
         Height          =   375
         Left            =   0
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "项目"
         CapAlign        =   2
         BackStyle       =   3
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   12632256
         Focus           =   0   'False
         cGradient       =   12632256
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "FrmResults.frx":C750
         cBack           =   16777215
      End
      Begin DeepLook.ucButtons_H btnSaveUVSList 
         Height          =   375
         Left            =   7320
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   0
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         Caption         =   "保存未用变量列表"
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
         Image           =   "FrmResults.frx":CCA2
         cBack           =   16777215
      End
   End
   Begin DeepLook.ucDeepLookHeader hedDeepLookHeader 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   13275
      _ExtentX        =   23521
      _ExtentY        =   661
   End
   Begin DeepLook.ucButtons_H optStatistics 
      Height          =   315
      Left            =   1200
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   405
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Caption         =   "统计"
      CapAlign        =   2
      BackStyle       =   1
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   12632256
      Focus           =   0   'False
      cGradient       =   12632256
      Gradient        =   3
      Mode            =   2
      Value           =   -1  'True
      Image           =   "FrmResults.frx":D1F6
      cBack           =   16777215
   End
   Begin DeepLook.ucButtons_H optUnusedVariables 
      Height          =   315
      Left            =   2085
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   405
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      Caption         =   "未用变量"
      CapAlign        =   2
      BackStyle       =   1
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   12632256
      Focus           =   0   'False
      cGradient       =   12632256
      Gradient        =   3
      Mode            =   2
      Value           =   0   'False
      Image           =   "FrmResults.frx":D4B8
      cBack           =   16777215
   End
   Begin MSComctlLib.TreeView TreeView 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   9763
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "ilstImages"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fmePieChart 
      BackColor       =   &H00D5E6EA&
      Height          =   5595
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   12615
      Begin VB.PictureBox pbxPieChartSubFrame 
         BackColor       =   &H00D5E6EA&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   120
         ScaleHeight     =   2655
         ScaleWidth      =   11775
         TabIndex        =   16
         Top             =   600
         Width           =   11775
         Begin MSChart20Lib.MSChart chtGroupChart 
            Height          =   2175
            Left            =   0
            OleObjectBlob   =   "FrmResults.frx":D77A
            TabIndex        =   28
            Top             =   -480
            Width           =   3015
         End
         Begin MSChart20Lib.MSChart chtFileChart 
            Height          =   2175
            Left            =   6240
            OleObjectBlob   =   "FrmResults.frx":F4CB
            TabIndex        =   30
            Top             =   -480
            Width           =   2535
         End
         Begin MSChart20Lib.MSChart chtProjectChart 
            Height          =   2175
            Left            =   2880
            OleObjectBlob   =   "FrmResults.frx":11873
            TabIndex        =   50
            Top             =   -480
            Width           =   3015
         End
         Begin MSChart20Lib.MSChart chtSPFChart 
            Height          =   2175
            Left            =   9000
            OleObjectBlob   =   "FrmResults.frx":13906
            TabIndex        =   51
            Top             =   -480
            Width           =   3015
         End
         Begin VB.Label lblSelSPFName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   9360
            TabIndex        =   37
            Top             =   1920
            Width           =   2415
         End
         Begin VB.Label lblSelectedSPFHeader 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "所选 S/P/F"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9360
            TabIndex        =   36
            Top             =   1680
            Width           =   2295
         End
         Begin VB.Shape shpKeyColour 
            BackColor       =   &H00C0C0C0&
            FillColor       =   &H00CC734D&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   19
            Left            =   10680
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label lblSPFBlankLines 
            BackStyle       =   0  'Transparent
            Caption         =   "空白行数"
            Height          =   270
            Left            =   10920
            TabIndex        =   35
            Top             =   2160
            Width           =   735
         End
         Begin VB.Shape shpKeyColour 
            BackColor       =   &H00C0C0C0&
            FillColor       =   &H000080FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   18
            Left            =   9360
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label lblSPFCodeLines 
            BackStyle       =   0  'Transparent
            Caption         =   "代码行数"
            Height          =   225
            Left            =   9615
            TabIndex        =   34
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label lblSelFileName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   6360
            TabIndex        =   32
            Top             =   1920
            Width           =   2295
         End
         Begin VB.Label lblSelProjName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   3240
            TabIndex        =   31
            Top             =   1920
            Width           =   2295
         End
         Begin VB.Label lblFileCodeLines 
            BackStyle       =   0  'Transparent
            Caption         =   "代码行数"
            Height          =   210
            Left            =   6600
            TabIndex        =   27
            Top             =   2160
            Width           =   855
         End
         Begin VB.Shape shpKeyColour 
            BackColor       =   &H00C0C0C0&
            FillColor       =   &H000080FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   17
            Left            =   6360
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label lblFileBlankLines 
            BackStyle       =   0  'Transparent
            Caption         =   "空白行数"
            Height          =   240
            Left            =   7920
            TabIndex        =   26
            Top             =   2160
            Width           =   735
         End
         Begin VB.Shape shpKeyColour 
            BackColor       =   &H00C0C0C0&
            FillColor       =   &H00CC734D&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   16
            Left            =   7680
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label lblFileCommentLines 
            BackStyle       =   0  'Transparent
            Caption         =   "注释行数"
            Height          =   210
            Left            =   7080
            TabIndex        =   25
            Top             =   2400
            Width           =   1095
         End
         Begin VB.Shape shpKeyColour 
            BackColor       =   &H00C0C0C0&
            FillColor       =   &H0000C000&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   15
            Left            =   6840
            Top             =   2400
            Width           =   135
         End
         Begin VB.Shape shpKeyColour 
            BackColor       =   &H00C0C0C0&
            FillColor       =   &H0000C000&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   14
            Left            =   3720
            Top             =   2400
            Width           =   135
         End
         Begin VB.Label lblPrjCommentLines 
            BackStyle       =   0  'Transparent
            Caption         =   "注释行数"
            Height          =   180
            Left            =   3960
            TabIndex        =   24
            Top             =   2400
            Width           =   1095
         End
         Begin VB.Shape shpKeyColour 
            BackColor       =   &H00C0C0C0&
            FillColor       =   &H00CC734D&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   13
            Left            =   4560
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label lblPrjBlankLines 
            BackStyle       =   0  'Transparent
            Caption         =   "空白行数"
            Height          =   165
            Left            =   4800
            TabIndex        =   23
            Top             =   2160
            Width           =   735
         End
         Begin VB.Shape shpKeyColour 
            BackColor       =   &H00C0C0C0&
            FillColor       =   &H000080FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   12
            Left            =   3240
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label lblPrjCodeLines 
            BackStyle       =   0  'Transparent
            Caption         =   "代码行数"
            Height          =   195
            Left            =   3480
            TabIndex        =   22
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label lblProjectGroupHeading 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "工程组"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   1680
            Width           =   2295
         End
         Begin VB.Label lblSelectedProjectHeader 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "所选工程"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3240
            TabIndex        =   20
            Top             =   1680
            Width           =   2295
         End
         Begin VB.Label lblSelectedFileHeader 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "所选文件"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6360
            TabIndex        =   19
            Top             =   1680
            Width           =   2295
         End
         Begin VB.Label lblGrpCodeLines 
            BackStyle       =   0  'Transparent
            Caption         =   "代码行数"
            Height          =   180
            Left            =   600
            TabIndex        =   18
            Top             =   2160
            Width           =   855
         End
         Begin VB.Shape shpKeyColour 
            BackColor       =   &H00C0C0C0&
            FillColor       =   &H000080FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   10
            Left            =   360
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label lblGrpBlankLines 
            BackStyle       =   0  'Transparent
            Caption         =   "空白行数"
            Height          =   240
            Left            =   1920
            TabIndex        =   17
            Top             =   2160
            Width           =   735
         End
         Begin VB.Shape shpKeyColour 
            BackColor       =   &H00C0C0C0&
            FillColor       =   &H00CC734D&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   11
            Left            =   1680
            Top             =   2160
            Width           =   135
         End
      End
   End
   Begin MSComctlLib.ListView lstVarList 
      Height          =   5535
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   9763
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "ilstUnusedVarImages"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "工程文件"
         Object.Width           =   4939
         ImageKey        =   "Project"
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "文件"
         Object.Width           =   4939
         ImageKey        =   "File"
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "类型/SPF 位置"
         Object.Width           =   5292
         ImageKey        =   "GlobalLocal"
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "变量名称"
         Object.Width           =   4939
         ImageKey        =   "Name"
      EndProperty
   End
   Begin VB.Label lblExternal 
      BackStyle       =   0  'Transparent
      Caption         =   "外部"
      Height          =   195
      Left            =   5145
      TabIndex        =   14
      Top             =   465
      Width           =   615
   End
   Begin VB.Shape shpKeyColour 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   4920
      Top             =   480
      Width           =   135
   End
   Begin VB.Label lblPrivate 
      BackStyle       =   0  'Transparent
      Caption         =   "私有变量"
      Height          =   195
      Left            =   5745
      TabIndex        =   13
      Top             =   465
      Width           =   735
   End
   Begin VB.Shape shpKeyColour 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   5580
      Top             =   480
      Width           =   135
   End
   Begin VB.Label lblPublic 
      BackStyle       =   0  'Transparent
      Caption         =   "全局变量"
      Height          =   195
      Left            =   6705
      TabIndex        =   12
      Top             =   465
      Width           =   795
   End
   Begin VB.Shape shpKeyColour 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   6555
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape shpKeyColour 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   7545
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape shpKeyColour 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   8340
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape shpKeyColour 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   9060
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape shpKeyColour 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   7
      Left            =   12480
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape shpKeyColour 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   8
      Left            =   10800
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape shpKeyColour 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   9
      Left            =   11640
      Top             =   480
      Width           =   135
   End
   Begin VB.Label lblEmpty 
      BackStyle       =   0  'Transparent
      Caption         =   "空"
      Height          =   255
      Left            =   10245
      TabIndex        =   5
      Top             =   465
      Width           =   615
   End
   Begin VB.Shape shpKeyColour 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   6
      Left            =   10065
      Top             =   480
      Width           =   135
   End
   Begin VB.Label lblScanResults 
      BackStyle       =   0  'Transparent
      Caption         =   "扫描分析结果:"
      Height          =   255
      Left            =   45
      TabIndex        =   1
      Top             =   450
      Width           =   1215
   End
   Begin VB.Label lblNormal 
      BackStyle       =   0  'Transparent
      Caption         =   "常规"
      Height          =   255
      Left            =   7755
      TabIndex        =   11
      Top             =   465
      Width           =   855
   End
   Begin VB.Label lblFriend 
      BackStyle       =   0  'Transparent
      Caption         =   "友元"
      Height          =   210
      Left            =   8520
      TabIndex        =   10
      Top             =   465
      Width           =   615
   End
   Begin VB.Label lblStatic 
      BackStyle       =   0  'Transparent
      Caption         =   "静态变量"
      Height          =   180
      Left            =   9255
      TabIndex        =   9
      Top             =   465
      Width           =   720
   End
   Begin VB.Label lblLet 
      BackStyle       =   0  'Transparent
      Caption         =   "Let"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   12720
      TabIndex        =   8
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblSet 
      BackStyle       =   0  'Transparent
      Caption         =   "Set"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   11880
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblGet 
      BackStyle       =   0  'Transparent
      Caption         =   "Get"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   11040
      TabIndex        =   7
      Top             =   480
      Width           =   615
   End
   Begin VB.Menu mnuShow 
      Caption         =   "Show"
      Visible         =   0   'False
      Begin VB.Menu mnuShowSUBS 
         Caption         =   "过程"
      End
      Begin VB.Menu mnuShowFUNCTIONS 
         Caption         =   "函数"
      End
      Begin VB.Menu mnuShowPROPERTIES 
         Caption         =   "属性"
      End
      Begin VB.Menu mnuShowEVENTS 
         Caption         =   "事件"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowFORMS 
         Caption         =   "窗体"
      End
      Begin VB.Menu mnuShowMODULES 
         Caption         =   "模块"
      End
      Begin VB.Menu mnuShowCLASSES 
         Caption         =   "类模块"
      End
      Begin VB.Menu mnuShowUC 
         Caption         =   "用户控件"
      End
      Begin VB.Menu mnuShowUD 
         Caption         =   "用户文档"
      End
      Begin VB.Menu mnuShowPP 
         Caption         =   "属性页"
      End
      Begin VB.Menu mnuShowDS 
         Caption         =   "设计器"
      End
      Begin VB.Menu mnuShowAllVB 
         Caption         =   "全部VB项目"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowREFCOM 
         Caption         =   "引用,组件和声明DLL"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowall 
         Caption         =   "全部"
      End
   End
   Begin VB.Menu mnuNetshow 
      Caption         =   "NETshow"
      Visible         =   0   'False
      Begin VB.Menu mnuShowIMPORTS 
         Caption         =   "导入"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowall2 
         Caption         =   "全部"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "ExOptns"
      Visible         =   0   'False
      Begin VB.Menu mnuHLGHTpmc 
         Caption         =   "高亮显示潜在的恶意代码"
      End
      Begin VB.Menu mnuHLGHTespfs 
         Caption         =   "高亮显示空 SPF's"
      End
      Begin VB.Menu mnuHLGHTexsf 
         Caption         =   "高亮显示扩展 SF's"
      End
   End
End
Attribute VB_Name = "FrmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:10:48
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：FrmResults
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:10:48
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V4.12.0
'*************************************************************************
Option Explicit

' -----------------------------------------------------------------------------------------------
Private CurrSelProject As String, CurrSelFile As String, CurrSelSPF As String
' -----------------------------------------------------------------------------------------------

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then DragForm Me.hwnd ' You can move the forms by holding down the left mouse button
End Sub

Private Sub mnuShowDS_Click()
    ExpandKeyByPic "Designer"
     
End Sub

Private Sub MNUshowIMPORTS_Click()
    ExpandKeyByPic "SysDLL"
    ExpandKey "NETfiles"
End Sub

Private Sub MNUshowMODULES_Click()
    ExpandKeyByPic "Module"
End Sub

Private Sub MNUshowPP_Click()
    ExpandKeyByPic "PropertyPage"
End Sub

Private Sub MNUshowPROPERTIES_Click()
    ExpandKey "_PROPERTIES"
End Sub

Private Sub MNUshowREFCOM_Click()
    ExpandKeyByPic "SysDLL", True
    ExpandKeyByPic "DLL", True
    ExpandKeyByPic "Component", True
    ExpandKeyByPic "CreateObject", True
End Sub

Private Sub MNUshowSUBS_Click()
    ExpandKey "_SUBS"
End Sub

Private Sub MNUshowUC_Click()
    ExpandKeyByPic "UserControl"
End Sub

Private Sub MNUshowUD_Click()
    ExpandKeyByPic "UserDocument"
End Sub

Private Sub optStatistics_Click()
    ShowHideDisplayElements
End Sub

Private Sub optUnusedVariables_Click()
    ShowHideDisplayElements
End Sub

Private Sub optCharts_Click()
    ShowHideDisplayElements
End Sub

Private Sub TreeView_DblClick()
    With TreeView
        If InStrB(1, .SelectedItem.Text, "(Double-Click", vbTextCompare) <> 0 Then
            Shell "Notepad.exe " & Mid$(.SelectedItem.Key, InStrRev(.SelectedItem.Key, "_") + 1), vbNormalFocus
        End If
    End With
End Sub

Private Sub TreeView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button > 1 Then PopupMenu mnuPopup
End Sub

Private Sub TreeView_KeyPress(KeyAscii As Integer)
    TreeView_Click
End Sub

Private Sub TreeView_NodeCheck(ByVal Node As MSComctlLib.Node)
    TreeView_Click
End Sub

Private Sub TreeView_NodeClick(ByVal Node As MSComctlLib.Node)
    TreeView_Click
End Sub

Private Sub TreeView_Validate(Cancel As Boolean)
    TreeView_Click
End Sub

Private Sub TreeView_Click()
    Dim CurrKey As String

    With TreeView
        CurrKey = Replace$(Left$(.SelectedItem.FullPath, InStrRev(.SelectedItem.FullPath, "\")), "\ ", "\")
        CurrKey = Replace$(CurrKey, "&", "&&", 1)

        If CurrKey = "" Then
            sbrStatus.Caption = "STAT>" & DefaultStatText
        Else
            If ShowItemKey = True Then
                sbrStatus.Caption = "STAT>" & .SelectedItem.Key & " (" & .SelectedItem.Index & ") 合计节点: " & .Nodes.Count
            Else
                sbrStatus.Caption = "STAT>当前选中节点: " & CurrKey
            End If
        End If

        On Error Resume Next
        CreateProjectPieGraph CurrKey ' Update the Selected Project pie graph data
        CreateFilePieGraph .SelectedItem.FullPath ' Update the Selected File pie graph data
        CreateSPFPieGraph ' Update the Selected Sub/Function/Property pie graph data
        On Error GoTo 0
    End With
End Sub

Private Sub btnAbout_Click()
    FrmAbout.Show 1
End Sub

Private Sub btnCollapseAll_Click()
    Dim I As Long
    I = SendMessage(TreeView.hwnd, WM_SETREDRAW, 0, 0)

    For I = 1 To TreeView.Nodes.Count
        TreeView.Nodes(I).Expanded = False
    Next

    TreeView.Nodes(1).Expanded = True

    I = SendMessage(TreeView.hwnd, WM_SETREDRAW, 1, 0)
End Sub

Private Sub btnExitDeepLook_Click()
    IsExit = True
    Unload Me
End Sub

Private Sub btnExpand_Click()
    If PROJECTMODE = VB6 Then PopupMenu mnuShow
    If PROJECTMODE = NET Then PopupMenu mnuNetshow
End Sub

Private Sub btnFileCopy_Click()
    With FrmCopyReport
        .lblCopyDir.Caption = GetRootDirectory(ProjectPath) & "Res\"
        .tvwItemsTV.Nodes.Clear
        .tvwNonCopyItemsTV.Nodes.Clear
        .btnStartCopy.Enabled = True
        .Caption = "复制所需文件的报告"
        .pgbPercentBar.Value = 0
        ModFileSearchHandler.LoadFiles
        .Visible = True '  \ A workaround; the form needs to be visible for the
        .CheckCheckboxes ' | the checkboxes in the treeview to check - the
        .Visible = False ' | form is made visible, the items checked before
        .Show 1 '          / it is hidden again and show modally
    End With
End Sub

Private Sub btnGenerateReport_Click()
    btnGenerateReport.Enabled = False
    btnGenerateReport.Caption = "稍候..."
        
    FrmReport.CreateTempXMLReport
    
    btnGenerateReport.Caption = "显示工程报告"
    btnGenerateReport.Enabled = True
    
    FrmReport.Show 1
End Sub

'保存未用变量列表
Private Sub btnSaveUVSList_Click()
    Dim I As Long, ProjFName As String * 30, FName As String * 30, SPFL As String * 30, VName As String * 30
    Dim FileNum As Integer
    
    cdgDialogs.Filter = "工程扫描报告文件 (*.txt)|*.txt"
    cdgDialogs.ShowSave

    FileNum = FreeFile
    
    If cdgDialogs.FileName = vbNullString Then Exit Sub

    If FileExists(cdgDialogs.FileName) Then
        I = MsgBoxEx("文件 """ & cdgDialogs.FileName & """ 已经存在. 是否覆盖?", vbExclamation Or vbYesNo, "提示", , , , , PicReport)
        If I = vbNo Then Exit Sub
    End If

    If LenB(cdgDialogs.FileName) Then
        Open cdgDialogs.FileName For Output As FileNum
        Print #FileNum, "------------------------------------"
        Print #FileNum, " DeepLook Unused Variable List File"
        Print #FileNum, "------------------------------------"
        Print #FileNum, "--   (C) Dean Camera, 2003-2005   --"
        Print #FileNum, "------------------------------------" & vbCrLf
        Print #FileNum, "File Created: " & Now & vbCrLf & vbCrLf
        Print #FileNum, "-------------------------------+--------------------------------+--------------------------------+-----------------------"
        Print #FileNum, " Project File Name:" & Space$(12) & "| Filename:" & Space$(22) & "| Type/SPF Location" & Space$(14) & "| Variable Name:"
        Print #FileNum, "-------------------------------+--------------------------------+--------------------------------+-----------------------"

        With lstVarList.ListItems
            On Local Error Resume Next
            For I = 1 To .Count - 2 ' Add each of the unused variables to the file (-2 to skip the "Total:" and blank list items)
                ProjFName = " " & Mid$(.Item(I).Text, 1, 30) & Space$(30 - Len(Mid$(.Item(I).Text, 1, 30)))
                FName = Mid$(.Item(I).ListSubItems(1).Text, 1, 30) & Space$(30 - Len(Mid$(.Item(I).ListSubItems(1).Text, 1, 30)))
                SPFL = Mid$(.Item(I).ListSubItems(2).Text, 1, 30) & Space$(30 - Len(Mid$(.Item(I).ListSubItems(2).Text, 1, 30)))
                VName = Mid$(.Item(I).ListSubItems(3).Text, 1, 30) & Space$(30 - Len(Mid$(.Item(I).ListSubItems(3).Text, 1, 30)))
                Print #FileNum, ProjFName & " | " & FName & " | "; SPFL & " | "; VName
            Next
            On Local Error GoTo 0

            Print #FileNum, "-------------------------------+--------------------------------+--------------------------------+-----------------------"
            Print #FileNum, vbCrLf & .Item(.Count).Text  ' Add the "Total:" node text to the file
        End With

        Close #FileNum
    End If
End Sub

Private Sub btnScanAnother_Click()
    FrmSelProject.Show
    Unload Me
End Sub

Private Sub Form_Load()
    IsExit = False

    sbrStatus.CaptionAlign = vbLeftJustify

    ' Colour the small key indicators:
    shpKeyColour(0).FillColor = RGB(150, 150, 150)
    shpKeyColour(1).FillColor = RGB(130, 0, 200)
    shpKeyColour(2).FillColor = RGB(200, 0, 150)
    shpKeyColour(3).FillColor = RGB(10, 150, 10)
    shpKeyColour(4).FillColor = RGB(20, 90, 100)
    shpKeyColour(5).FillColor = RGB(50, 23, 80)
    shpKeyColour(6).FillColor = RGB(255, 0, 0)
    shpKeyColour(7).FillColor = RGB(249, 164, 0)
    shpKeyColour(8).FillColor = RGB(217, 206, 19)
    shpKeyColour(9).FillColor = RGB(19, 217, 192)

    ShowHideDisplayElements ' Make sure the treeview is the only display visible

    RemoveTabStops Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then ' "X" button or ALT+F4 pressed
        IsExit = True
        KillProgram False
    ElseIf IsExit = True Then
        KillProgram False
    End If
End Sub

Private Sub Form_Resize()
    On Local Error Resume Next

    If Me.Height < 5685 Then Me.Height = 5685: Exit Sub
    If Me.Width < 13270 Then Me.Width = 13270: Exit Sub

    sbrStatus.Width = Me.Width - 125
    sbrStatus.Top = Me.ScaleHeight - sbrStatus.Height

    pbxControlPanel.Top = Me.Height - pbxControlPanel.Height - 850
    pbxControlPanel.Left = (Me.Width / 2) - (pbxControlPanel.Width / 2)

    TreeView.Width = Me.Width - 350
    lstVarList.Width = TreeView.Width
    fmePieChart.Width = TreeView.Width

    TreeView.Height = pbxControlPanel.Top - 850
    lstVarList.Height = TreeView.Height
    fmePieChart.Height = TreeView.Height

    pbxPieChartSubFrame.Top = fmePieChart.Top + (fmePieChart.Height / 2) - (pbxPieChartSubFrame.Height / 1.5)
    pbxPieChartSubFrame.Left = fmePieChart.Left + (fmePieChart.Width / 2) - (pbxPieChartSubFrame.Width / 2)

    btnAbout.Left = Me.Width - btnAbout.Width - 110
    hedDeepLookHeader.ResizeMe
End Sub

Private Sub MNUHLGHTespfs_Click()
    Dim I As Long

    mnuHLGHTespfs.Checked = Not mnuHLGHTespfs.Checked

    SendMessage TreeView.hwnd, WM_SETREDRAW, 0, 0 ' Stop the treeview from redrawing

    If mnuHLGHTespfs.Checked Then
        For I = 1 To TreeView.Nodes.Count
            If TreeView.Nodes(I).ForeColor = RGB(255, 0, 0) Then
                TreeView.Nodes(I).ForeColor = RGB(255, 255, 255)
                TreeView.Nodes(I).BackColor = RGB(255, 0, 0)
            End If
        Next
    Else
        For I = 1 To TreeView.Nodes.Count
            If TreeView.Nodes(I).BackColor = RGB(255, 0, 0) Then
                TreeView.Nodes(I).BackColor = RGB(255, 255, 255)
                TreeView.Nodes(I).ForeColor = RGB(255, 0, 0)
            End If
        Next
    End If

    SendMessage TreeView.hwnd, WM_SETREDRAW, 1, 0 ' Reenable treeview drawing
End Sub

Private Sub MNUHLGHTexsf_Click()
    Dim I As Long

    mnuHLGHTexsf.Checked = Not mnuHLGHTexsf.Checked

    SendMessage TreeView.hwnd, WM_SETREDRAW, 0, 0

    If mnuHLGHTexsf.Checked Then
        For I = 1 To TreeView.Nodes.Count
            If TreeView.Nodes(I).ForeColor = RGB(150, 150, 150) Then
                TreeView.Nodes(I).ForeColor = RGB(255, 255, 255)
                TreeView.Nodes(I).BackColor = RGB(200, 200, 200)
            End If
        Next
    Else
        For I = 1 To TreeView.Nodes.Count
            If TreeView.Nodes(I).BackColor = RGB(200, 200, 200) Then
                TreeView.Nodes(I).BackColor = RGB(255, 255, 255)
                TreeView.Nodes(I).ForeColor = RGB(150, 150, 150)
            End If
        Next
    End If

    SendMessage TreeView.hwnd, WM_SETREDRAW, 1, 0
End Sub

Private Sub MNUHLGHTpmc_Click()
    Dim I As Long

    mnuHLGHTpmc.Checked = Not mnuHLGHTpmc.Checked

    SendMessage TreeView.hwnd, WM_SETREDRAW, 0, 0

    If mnuHLGHTpmc.Checked Then
        For I = 1 To TreeView.Nodes.Count
            If InStrB(1, TreeView.Nodes(I).Key, "_PMC_") <> 0 Then
                If Right$(TreeView.Nodes(I).Key, 5) <> "COUNT" Then
                    TreeView.Nodes(I).ForeColor = RGB(255, 255, 255)
                    TreeView.Nodes(I).BackColor = RGB(0, 180, 0)
                End If
            End If
        Next
    Else
        For I = 1 To TreeView.Nodes.Count
            If InStrB(1, TreeView.Nodes(I).Key, "_PMC_") <> 0 Then
                If Right$(TreeView.Nodes(I).Key, 5) <> "COUNT" Then
                    TreeView.Nodes(I).BackColor = RGB(255, 255, 255)
                    TreeView.Nodes(I).ForeColor = RGB(0, 0, 0)
                End If
            End If
        Next
    End If

    SendMessage TreeView.hwnd, WM_SETREDRAW, 1, 0
End Sub

Private Sub MNUshowAll_Click()
    Dim I As Long
    I = SendMessage(TreeView.hwnd, WM_SETREDRAW, 0, 0)

    For I = 1 To TreeView.Nodes.Count
        TreeView.Nodes(I).Expanded = True
    Next

    I = SendMessage(TreeView.hwnd, WM_SETREDRAW, 1, 0)
End Sub

Private Sub MNUshowall2_Click()
    MNUshowAll_Click
End Sub

Private Sub MNUshowAllVB_Click()
    ExpandKeyByPic "Form"
    ExpandKeyByPic "Class"
    ExpandKeyByPic "Module"
    ExpandKeyByPic "PropertyPage"
    ExpandKeyByPic "UserControl"
    ExpandKeyByPic "UserDocument"
    ExpandKeyByPic "Designer"
End Sub

Private Sub MNUshowCLASSES_Click()
    ExpandKeyByPic "Class"
End Sub

Private Sub MNUshowEVENTS_Click()
    ExpandKey "_EVENTS"
End Sub

Private Sub MNUshowFORMS_Click()
    ExpandKeyByPic "Form"
End Sub

Private Sub MNUshowFUNCTIONS_Click()
    ExpandKey "_FUNCTIONS"
End Sub

Private Sub ExpandKey(Suffix As String)
    Dim I As Long, X As Long, Z As Long
    X = SendMessage(TreeView.hwnd, WM_SETREDRAW, 0, 0)

    On Local Error Resume Next
    
    For I = 1 To TreeView.Nodes.Count
        If Right$(TreeView.Nodes(I).Key, Len(Suffix)) = Suffix Then
            TreeView.Nodes(I).Expanded = True
            TreeView.Nodes(I).Parent.Expanded = True
            Z = TreeView.Nodes(I).Parent.Index
            TreeView.Nodes(Z).Parent.Expanded = True
            TreeView.Nodes(TreeView.Nodes(Z).Parent.Index).Parent.Expanded = True
        End If

        If TreeView.Nodes(I).Key = "GROUP" Then TreeView.Nodes(I).Expanded = True
    Next

    X = SendMessage(TreeView.hwnd, WM_SETREDRAW, 1, 0)
End Sub

Private Sub ExpandKeyByPic(ImageKey As String, Optional DoubleParent As Boolean)
    Dim I As Long, Z As Long
    I = SendMessage(TreeView.hwnd, WM_SETREDRAW, 0, 0)

    On Local Error Resume Next

    For I = 1 To TreeView.Nodes.Count
        If TreeView.Nodes(I).Image = ImageKey Then
            TreeView.Nodes(I).Expanded = True
            TreeView.Nodes(I).Parent.Expanded = True

            If DoubleParent = True Then
                Z = TreeView.Nodes(I).Parent.Index
                TreeView.Nodes(Z).Parent.Expanded = True
            End If
        End If

        If TreeView.Nodes(I).Key = "GROUP" Then TreeView.Nodes(I).Expanded = True
    Next

    I = SendMessage(TreeView.hwnd, WM_SETREDRAW, 1, 0)
End Sub

Private Sub CreateProjectPieGraph(CurrKey As String)
    Dim SLoc As Integer, TempKey As String, Eloc As Long
    Dim PLines As Long, PLinesNB As Long, PLinesC As Long, BlankLines As Long

    With TreeView
        If ((LenB(CurrKey) And CurrKey <> "Project Group\") Or .SelectedItem.Image = "Project") And .SelectedItem.Image <> "Unknown" Then
            CurrKey = CurrKey & .SelectedItem.Text

            If Left$(.SelectedItem.FullPath, 14) = "Project Group\" Then
                CurrKey = Mid$(CurrKey, 15)
                If InStrB(1, CurrKey, "\") > 0 Then CurrKey = Left$(CurrKey, InStr(1, CurrKey, "\") - 1)

                lblSelProjName.Caption = Trim$(CurrKey)

                CurrKey = "_" & CurrKey
            Else
                CurrKey = "PROJECT_?"

                lblSelProjName.Caption = Trim$(.SelectedItem.Root.Text)

                If lblSelProjName.Caption = "(Temp Project)" Then
                    lblSelProjName.Caption = vbNullString
                    Exit Sub
                End If
            End If

            If CurrKey <> CurrSelProject Then
                CurrSelProject = CurrKey

                TempKey = .Nodes(CurrKey & "_LINES").Text
                Eloc = InStr(1, TempKey, "[") - 2
                SLoc = InStrRev(TempKey, " ", Eloc) + 1
                TempKey = Mid$(TempKey, SLoc, Eloc - SLoc + 1)
                PLines = Int(Replace(TempKey, ",", ""))

                TempKey = .Nodes(CurrKey & "_LINESNB").Text
                Eloc = InStr(1, TempKey, "[") - 2
                SLoc = InStrRev(TempKey, " ", Eloc) + 1
                TempKey = Mid$(TempKey, SLoc, Eloc - SLoc + 1)
                PLinesNB = Int(Replace(TempKey, ",", ""))

                TempKey = .Nodes(CurrKey & "_LINESCOMMENT").Text
                Eloc = InStr(1, TempKey, "[") - 2
                SLoc = InStrRev(TempKey, " ", Eloc) + 1
                TempKey = Mid$(TempKey, SLoc, Eloc - SLoc + 1)
                PLinesC = Int(Replace(TempKey, ",", ""))

                BlankLines = (PLines - PLinesNB)
                With chtProjectChart
                    .Column = 1
                    .Data = Round((100 / OneIfNull(PLines + BlankLines + PLinesC)) * OneIfNull(PLines))
                    .Column = 2
                    .Data = Round((100 / OneIfNull(PLines + BlankLines + PLinesC)) * OneIfNull(BlankLines))
                    .Column = 3
                    .Data = Round((100 / OneIfNull(PLines + BlankLines + PLinesC)) * OneIfNull(PLinesC))
                End With
            End If
        Else
            For SLoc = 1 To 3
                Me.chtProjectChart.Column = SLoc
                Me.chtProjectChart.Data = 0
            Next

            lblSelProjName.Caption = vbNullString
            CurrSelProject = vbNullString
        End If
    End With
End Sub

Private Sub CreateFilePieGraph(CurrKey As String)
    Dim TempKey As String, PLines As Long, PLinesNB As Long, PLinesC As Long, BlankLines As Long
    Dim Spos As Long, Epos As Long

    If InStrB(1, CurrKey, "Forms\") Then
        Spos = InStr(1, CurrKey, "Forms\") + 6
    ElseIf InStrB(1, CurrKey, "Modules\") Then
        Spos = InStr(1, CurrKey, "Modules\") + 8
    ElseIf InStrB(1, CurrKey, "Classes\") Then
        Spos = InStr(1, CurrKey, "Classes\") + 8
    ElseIf InStrB(1, CurrKey, "User Controls\") Then
        Spos = InStr(1, CurrKey, "User Controls\") + 14
    ElseIf InStrB(1, CurrKey, "User Documents\") Then
        Spos = InStr(1, CurrKey, "User Documents\") + 15
    ElseIf InStrB(1, CurrKey, "Designers\") Then
        Spos = InStr(1, CurrKey, "Designers\") + 10
    ElseIf InStrB(1, CurrKey, "Property Pages\") Then
        Spos = InStr(1, CurrKey, "Property Pages\") + 15
    Else
        GoTo NotFileNode
    End If

    Epos = InStr(Spos, CurrKey, "\")
    If Epos Then CurrKey = Left$(CurrKey, Epos - 1)

    lblSelFileName.Caption = Trim$(Mid$(CurrKey, Spos))

    On Local Error GoTo NotFileNode
    CurrKey = Replace$(CurrKey, "\", "_")
    If TreeView.Nodes(1).Text = "Project Group" Then
        CurrKey = "_" & Mid$(CurrKey, InStr(1, CurrKey, "_") + 1)
    Else
        CurrKey = "PROJECT_?_" & Mid$(CurrKey, InStr(1, CurrKey, "_") + 1)
    End If

    CurrKey = Replace$(CurrKey, "Forms", "FORMS")
    CurrKey = Replace$(CurrKey, "Modules", "MODULES")
    CurrKey = Replace$(CurrKey, "Classes", "CLASSES")
    CurrKey = Replace$(CurrKey, "User Controls", "USERCONTROLS")
    CurrKey = Replace$(CurrKey, "User Documents", "USERDOCUMENTS")
    CurrKey = Replace$(CurrKey, "Designers", "DESIGNERS")
    CurrKey = Replace$(CurrKey, "Property Pages", "PROPERTYPAGES")

    TempKey = CurrKey & "_LINESNB"
    TempKey = Replace$(TreeView.Nodes(TempKey).Text, ",", "")
    TempKey = Mid$(TempKey, InStr(1, TempKey, ":") + 2)
    TempKey = Left$(TempKey, InStr(1, TempKey, " ") - 1)
    PLinesNB = Int(TempKey)

    TempKey = CurrKey & "_LINES"
    TempKey = Replace$(TreeView.Nodes(TempKey).Text, ",", "")
    TempKey = Mid$(TempKey, InStr(1, TempKey, ":") + 2)
    TempKey = Left$(TempKey, InStr(1, TempKey, " ") - 1)
    PLines = Int(TempKey)

    TempKey = CurrKey & "_COMMENTLINES"
    TempKey = Replace$(TreeView.Nodes(TempKey).Text, ",", "")
    TempKey = Mid$(TempKey, InStr(1, TempKey, ":") + 2)
    TempKey = Left$(TempKey, InStr(1, TempKey, " ") - 1)
    PLinesC = Int(TempKey)

    If CurrKey = CurrSelFile Then Exit Sub
    CurrSelFile = CurrKey

    BlankLines = (PLines - PLinesNB)

    With chtFileChart
        .Column = 1
        .Data = Round((100 / OneIfNull(PLines + BlankLines + PLinesC)) * OneIfNull(PLines))
        .Column = 2
        .Data = Round((100 / OneIfNull(PLines + BlankLines + PLinesC)) * OneIfNull(BlankLines))
        .Column = 3
        .Data = Round((100 / OneIfNull(PLines + BlankLines + PLinesC)) * OneIfNull(PLinesC))
    End With

    Exit Sub
NotFileNode:

    For Spos = 1 To 3
        Me.chtFileChart.Column = Spos
        Me.chtFileChart.Data = 0
    Next

    lblSelFileName.Caption = vbNullString
    CurrSelFile = vbNullString
End Sub

Private Sub CreateSPFPieGraph()
    Dim TempName As String, TempKey As String, SLoc As Long, NodeIndex As Long
    Dim PLines As Long, PLinesNB As Long

    NodeIndex = TreeView.SelectedItem.Index
    If TreeView.SelectedItem.Image = "Info" Then NodeIndex = TreeView.SelectedItem.Parent.Index

    With TreeView.Nodes(NodeIndex)
        TempName = Mid$(.Key, InStrRev(.Key, "_") + 1)

        If Left$(TempName, 3) = "SUB" Then
            SLoc = 4
        ElseIf Left$(TempName, 8) = "FUNCTION" Then
            SLoc = 9
        ElseIf Left$(TempName, 8) = "PROPERTY" Then
            SLoc = 9
        Else
            GoTo NotSPFNode
        End If

        If InStrB(1, .Text, " Lib ") Then GoTo NotSPFNode
        If Mid$(TempName, SLoc, 1) = "S" Then GoTo NotSPFNode
        If .Text = CurrSelSPF Then Exit Sub

        If InStrB(1, .Text, "(") Then
            lblSelSPFName.Caption = Left$(.Text, InStr(1, .Text, "(") - 1)
        Else
            lblSelSPFName.Caption = .Text
        End If

        CurrSelSPF = .Text

        TempKey = Replace$(TreeView.Nodes(.Key & "_LINES").Text, ",", "")
        TempKey = Mid$(TempKey, InStr(1, TempKey, ":") + 2)
        PLines = Int(Trim$(TempKey))

        TempKey = Replace$(TreeView.Nodes(.Key & "_LINESNB").Text, ",", "")
        TempKey = Mid$(TempKey, InStr(1, TempKey, ":") + 2)
        PLinesNB = Int(Trim$(TempKey))

        Me.chtSPFChart.Column = 1
        Me.chtSPFChart.Data = PLines
        Me.chtSPFChart.Column = 2
        Me.chtSPFChart.Data = PLines - PLinesNB
    End With

    Exit Sub
NotSPFNode:

    For SLoc = 1 To 2
        Me.chtSPFChart.Column = SLoc
        Me.chtSPFChart.Data = 0
    Next

    lblSelSPFName.Caption = vbNullString
    CurrSelSPF = vbNullString
End Sub

Private Sub ShowHideDisplayElements() ' Show/hide the display elements
    fmePieChart.Visible = optCharts.Value
    lstVarList.Visible = optUnusedVariables.Value
    TreeView.Visible = optStatistics.Value
End Sub
