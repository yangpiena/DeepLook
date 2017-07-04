VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsNETproject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:12:22
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：ClsNETproject
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:12:22
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V4.12.0
'*************************************************************************


Option Explicit

' -----------------------------------------------------------------------------------------------
Dim Code_LinesNB As Long
Dim Code_Lines As Long
Dim Blank_Lines As Long
Dim Comment_Lines As Long
Dim Total_Lines As Long
' -----------------------------------------------------------------------------------------------

Property Get TotalLines() As Long
    TotalLines = Total_Lines
End Property

Property Let TotalLines(AddAmount As Long)
    Total_Lines = Total_Lines + AddAmount
End Property

Property Get CodeLinesNB() As Long
    CodeLinesNB = Code_LinesNB
End Property

Property Let CodeLinesNB(AddAmount As Long)
    Code_LinesNB = Code_LinesNB + AddAmount
End Property

Property Get CodeLines() As Long
    CodeLines = Code_Lines
End Property

Property Let CodeLines(AddAmount As Long)
    Code_Lines = Code_Lines + AddAmount
End Property

Property Get BlankLines() As Long
    BlankLines = Blank_Lines
End Property

Property Let BlankLines(AddAmount As Long)
    Blank_Lines = Blank_Lines + AddAmount
End Property

Property Get CommentLines() As Long
    CommentLines = Comment_Lines
End Property

Property Let CommentLines(AddAmount As Long)
    Comment_Lines = Comment_Lines + AddAmount
End Property

