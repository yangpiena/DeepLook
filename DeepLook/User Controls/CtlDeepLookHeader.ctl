VERSION 5.00
Begin VB.UserControl ucDeepLookHeader 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5355
   Picture         =   "CtlDeepLookHeader.ctx":0000
   ScaleHeight     =   450
   ScaleWidth      =   5355
   ToolboxBitmap   =   "CtlDeepLookHeader.ctx":4B12
End
Attribute VB_Name = "ucDeepLookHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:13:00
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：ucDeepLookHeader
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:13:00
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V4.12.0
'*************************************************************************
Option Explicit


' Used to minimize memory requirements for DeepLook by only storing one logo

Private Sub UserControl_Resize()
    On Error Resume Next
    
    UserControl.Height = 370
    UserControl.Width = UserControl.Parent.Width
End Sub

Sub ResizeMe()
    UserControl_Resize
End Sub
