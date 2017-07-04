VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsXMLStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:12:47
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：ClsXMLStack
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:12:47
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V4.12.0
'*************************************************************************

' This class is used by the ModXMLReport module. I decided to take a leaf out of my microcontroller book
' and "borrow" the idea of using a psudo-stack to hold the XML headers because they are added and removed
' in a linear fashion, similar to a assembler stack. Unlike a true stack however, these functions also add
' the opening and closing headers to the XML file. The correct header is chosen via a enum in the Push method.
' Data and headers are indented with spaces according to the current stack depth.

' PUSH:       Pushes the parsed header onto the stack and adds the header to the XML file.
' PEEK:       Returns the topmost stack element without POPing it.
' OPENFILE:   Opens the XML file. This MUST be executed before the class is terminated.
' POP:        Pops the topmost element off the stack and adds the closing header to the XML file.
' PUSHADDPOP: Pushes a header, adds parsed data and pops the same header off the stack. used for some XML
'             headers which do not require multiple lines (ie. no children).
' ADD:        Adds data to the XML file without pushing a header onto the stack.
' POPALL:     Pops all headers off the stack sequentially (starting with the topmost element) and adds closing
'             headers to the XML file.
' STACKDEPTH: Returns the depth (number of elements) in the stack.

Option Explicit

' -----------------------------------------------------------------------------------------------
Private HeadStack() As String
Private HeadStackElements As Long

Private XMLFileNum As Integer
' -----------------------------------------------------------------------------------------------

Private Sub Class_Initialize()
    ReDim HeadStack(10) As String ' Create 10 blanks in the array
End Sub

Private Sub Class_Terminate()
    If XMLFileNum Then Close XMLFileNum
End Sub

' -----------------------------------------------------------------------------------------------

Public Sub OpenFile(FileName As String)
    XMLFileNum = FreeFile

    Open FileName For Output As #XMLFileNum
End Sub

Public Sub PushAddPop(hType As String, Data As String)
    Print #XMLFileNum, Space$(StackDepth + 2) & hType & Data & "</" & Mid$(hType, 2) ' Add data and header in a single line
End Sub

Public Sub Push(hType As String)
    Print #XMLFileNum, Space$(StackDepth + 2) & hType ' Print the header to the XML file

    HeadStackElements = StackDepth + 1
        
    If UBound(HeadStack) < HeadStackElements Then ' Not enough space left in the stack
        ReDim Preserve HeadStack(HeadStackElements + 10) As String ' Create another 10 blanks (faster to redim less frequently)
    End If
        
    HeadStack(HeadStackElements) = hType ' Add the data to the stack
End Sub

Public Sub Pop()
    Print #XMLFileNum, Space$(StackDepth + 1) & "</" & Mid$(Peek, 2) ' Print the closing header to the XML file

    HeadStackElements = HeadStackElements - 1 ' Pop the data off the stack
End Sub

Public Property Get Peek() As String
    Peek = HeadStack(StackDepth) ' Peek at the topmost data
End Property

Public Property Get StackDepth() As String
    StackDepth = HeadStackElements ' Get the stack depth
End Property

Public Sub PopAll()
    Dim I As Integer

    For I = 1 To StackDepth ' Pop all data off the stack
        Pop ' Pop the topmost data
    Next
    
    ReDim HeadStack(10) As String ' Clear the array to 10 blanks
End Sub

Public Sub Add(Data As String)
    Print #XMLFileNum, Data ' Print the data to the XML file (don't add it to the stack)
End Sub
