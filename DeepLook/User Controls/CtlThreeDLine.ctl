VERSION 5.00
Begin VB.UserControl ucThreeDLine 
   BackColor       =   &H00000000&
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   780
   ScaleHeight     =   330
   ScaleWidth      =   780
   ToolboxBitmap   =   "CtlThreeDLine.ctx":0000
End
Attribute VB_Name = "ucThreeDLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:15:58
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：ucThreeDLine
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:15:58
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V1.0.0
'*************************************************************************

' Simple 3D line control (auto orientation)

' This is my (Dean Camera) first attempt at a GDI control, so I decided to start VERY simple. No leaks though!
' For some reason will look odd in IDE when clicked, works fine when running.

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Public Enum LineType
 Solid = 0
 LONGDASH = 1
 SHORTDASH = 2
 DOTDASH = 3
 DOTDOTDASH = 4
End Enum

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private V_LineType As LineType
Private V_LineColour As OLE_COLOR

Public Property Get LineStyle() As LineType
    LineStyle = V_LineType

    UserControl_Paint
End Property

Public Property Let LineStyle(LType As LineType)
    V_LineType = LType

    UserControl_Paint
End Property

Public Property Get LineColour() As OLE_COLOR
    LineColour = V_LineColour

    UserControl_Paint
End Property

Public Property Let LineColour(Colour As OLE_COLOR)
    V_LineColour = Colour

    UserControl_Paint
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    V_LineType = PropBag.ReadProperty("LineStyle", 0)
    V_LineColour = PropBag.ReadProperty("LineColour", RGB(150, 150, 150))

    UserControl_Paint
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "LineStyle", V_LineType, 0
    PropBag.WriteProperty "LineColour", V_LineColour, RGB(150, 150, 150)

    UserControl_Paint
End Sub

Private Sub UserControl_Paint()
    Dim Pen As Long, SObj As Long
    Dim LColour As Long, pCoord As POINTAPI

    If Not UserControl.Parent Is Nothing Then
        UserControl.BackColor = UserControl.Parent.BackColor
    End If
    
    OleTranslateColor V_LineColour, &HFFFFFF, LColour
    MoveToEx hDC, 1, 1, pCoord
    
    If UserControl.Width > UserControl.Height Then
        UserControl.Height = 50

        Pen = CreatePen(V_LineType, 1, LColour)
        SObj = SelectObject(hDC, Pen)
        LineTo hDC, UserControl.Width / Screen.TwipsPerPixelX, 1
        SelectObject hDC, SObj
        DeleteObject Pen

        Pen = CreatePen(V_LineType, 1, vbWhite)
        SObj = SelectObject(hDC, Pen)
        MoveToEx hDC, 1, 2, pCoord
        LineTo hDC, UserControl.Width / Screen.TwipsPerPixelX, 2
        SelectObject hDC, SObj
        DeleteObject Pen
    Else
        UserControl.Width = 50

        Pen = CreatePen(V_LineType, 1, LColour)
        SObj = SelectObject(hDC, Pen)
        LineTo hDC, 1, UserControl.Height / Screen.TwipsPerPixelY
        SelectObject hDC, SObj
        DeleteObject Pen

        Pen = CreatePen(V_LineType, 1, vbWhite)
        SObj = SelectObject(hDC, Pen)
        MoveToEx hDC, 2, 1, pCoord
        LineTo hDC, 2, UserControl.Height / Screen.TwipsPerPixelY
        SelectObject hDC, SObj
        DeleteObject Pen
    End If
End Sub
