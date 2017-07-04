VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsEventLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:12:02
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：ClsEventLog
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:12:02
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V4.12.0
'*************************************************************************
' THIS MODULE WAS NOT WRITTEN BY DEAN CAMERA. I CANNOT OFFER ANY SUPPORT FOR THIS MODULE.


' Used in conjunction with MsgBoxEx.bas for custom message boxes.

Option Explicit

Private Const GMEM_ZEROINIT = &H40

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
    
Private Declare Function GlobalAlloc Lib "kernel32" _
    (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function GetLastError Lib "kernel32" () As Long

Private Declare Function RegisterEventSource Lib "advapi32.dll" Alias "RegisterEventSourceA" _
    (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
    
Private Declare Function DeregisterEventSource Lib "advapi32.dll" _
    (ByVal hEventLog As Long) As Long
    
Private Declare Function ReportEvent Lib "advapi32.dll" Alias "ReportEventA" _
    (ByVal hEventLog As Long, ByVal wType As Integer, ByVal wCategory As Integer, _
    ByVal dwEventID As Long, ByVal lpUserSid As Any, ByVal wNumStrings As Integer, _
    ByVal dwDataSize As Long, plpStrings As Long, lpRawData As Any) As Boolean
    
Public Enum EventLogType
    evnERROR = 1
    evnWARNING = 2
    evnINFORMATION = 4
End Enum

'local variable(s) to hold property value(s)
Private mvarApplication As String 'local copy
Private mvarEventID As String 'local copy
Private mvarDescription As String 'local copy
Private mvarLogType As EventLogType 'local copy
Private mvarServer As String 'local copy
Public Property Let Server(ByVal vData As String)
    mvarServer = vData
End Property
Public Property Get Server() As String
    Server = mvarServer
End Property
Public Property Let LogType(ByVal vData As EventLogType)
    mvarLogType = vData
End Property
Public Property Get LogType() As EventLogType
    LogType = mvarLogType
End Property
Public Property Let Description(ByVal vData As String)
    mvarDescription = vData
End Property
Public Property Get Description() As String
    Description = mvarDescription
End Property
Public Property Let EventID(ByVal vData As String)
    mvarEventID = vData
End Property
Public Property Get EventID() As String
    EventID = mvarEventID
End Property
Public Property Let Application(ByVal vData As String)
    mvarApplication = vData
End Property
Public Property Get Application() As String
    Application = mvarApplication
End Property
Public Function WriteEvent( _
        ByVal Description As String, _
        Optional ByVal Application As String = vbNullString, _
        Optional ByVal EventID As String = vbNullString, _
        Optional ByVal LogType As EventLogType _
    ) As Boolean
      
    Dim hEventLog As Long
    Dim hMsgs As Long
    Dim lngBufferSize As Long
    Dim OkToWrite As Boolean
    
    On Error Resume Next
    
    WriteEvent = False
    
    If LogType <> 0 Then mvarLogType = LogType
    If EventID <> vbNullString Then mvarEventID = EventID
    If Application <> vbNullString Then mvarApplication = Application
    If Description <> vbNullString Then mvarDescription = Description
    
    
    If (mvarLogType <> evnERROR) And _
        (mvarLogType <> evnWARNING) And _
        (mvarLogType <> evnINFORMATION) Then
        OkToWrite = False
    ElseIf mvarEventID <= 0 Then
        OkToWrite = False
    ElseIf mvarApplication = "" Then
        OkToWrite = False
    ElseIf mvarDescription = "" Then
        OkToWrite = False
    Else
        OkToWrite = True
    End If
    
    If OkToWrite Then
        If mvarServer = "" Then mvarServer = vbNullString
        lngBufferSize = Len(mvarDescription) + 1
        hMsgs = GlobalAlloc(GMEM_ZEROINIT, lngBufferSize)
        CopyMemory ByVal hMsgs, ByVal mvarDescription, lngBufferSize
        
        hEventLog = RegisterEventSource(mvarServer, mvarApplication)
        If ReportEvent(hEventLog, mvarLogType, 0, mvarEventID, 0&, 1, 0, hMsgs, 0&) Then
           WriteEvent = True
        End If
        DeregisterEventSource hEventLog
        
        GlobalFree hMsgs
    End If
    
End Function



