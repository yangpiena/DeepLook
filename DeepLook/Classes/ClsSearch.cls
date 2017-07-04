VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2026
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Files"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:12:41
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：ClsSearch
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:12:41
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V4.12.0
'*************************************************************************
' THIS MODULE WAS NOT WRITTEN BY DEAN CAMERA. I CANNOT OFFER ANY SUPPORT FOR THIS MODULE.


Option Explicit

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 400
    cAlternate As String * 14
End Type

Public Enum edwFileAttributes
    FILE_ATTRIBUTE_ARCHIVE = &H20
    FILE_ATTRIBUTE_COMPRESSED = &H800
    FILE_ATTRIBUTE_DIRECTORY = &H10
    FILE_ATTRIBUTE_HIDDEN = &H2
    FILE_ATTRIBUTE_NORMAL = &H80
    FILE_ATTRIBUTE_READONLY = &H1
    FILE_ATTRIBUTE_SYSTEM = &H4
    FILE_ATTRIBUTE_TEMPORARY = &H100
End Enum

Private Declare Function FindFirstFile _
        Lib "kernel32" _
        Alias "FindFirstFileA" ( _
        ByVal lpFileName As String, _
        lpFindFileData As WIN32_FIND_DATA) _
    As Long

Private Declare Function FindNextFile _
        Lib "kernel32" _
        Alias "FindNextFileA" ( _
        ByVal hFindFile As Long, _
        lpFindFileData As WIN32_FIND_DATA) _
    As Long

Private Declare Function FindClose _
        Lib "kernel32" ( _
        ByVal hFindFile As Long) _
    As Long
                            
Private Declare Function FileTimeToSystemTime _
        Lib "kernel32" ( _
        lpFileTime As FILETIME, _
        lpSystemTime As SYSTEMTIME) _
    As Long
                    
Private m_Files As ClsFiles

Public Property Get Files() As ClsFiles
Attribute Files.VB_Description = "Collection holding file search results."
Attribute Files.VB_HelpID = 2027
'##BLOCK_DESCRIPTION Collection holding file search results.
    Set Files = m_Files
End Property

Public Sub SearchFiles(ByVal FilePath As String, _
                       ByVal FileName As String, _
                       Optional ByVal Recursive As Boolean = False _
                       )
Attribute SearchFiles.VB_Description = "Call this function to search for file name."
Attribute SearchFiles.VB_HelpID = 2028
        '##PARAMETER_DESCRIPTION Recursive Set to true to include subfolders in search.
        '##PARAMETER_DESCRIPTION FileName Name of file to search for, wildcard allowed.
        '##PARAMETER_DESCRIPTION FilePath Windows path to begin file search.
        '##BLOCK_DESCRIPTION Call this function to search for file name.
       
        '<EhHeader>
        On Error GoTo SearchFiles_Error
        Dim strErrorDescription As String 'string to set extended error information
        'adding to debug call stack
        ModCallStack.StackAdd "Search", "SearchFiles", _
           "FilePath: " & FilePath & ", FileName: " & FileName & ", Recursive: " & Recursive, ""
        '</EhHeader>
    
        Dim typFindData As WIN32_FIND_DATA
        Dim lngFind As Long
        Dim lngMatch As Long
        Dim objFile As ClsFile
    
        'making sure the trailing backslash is at the end of the path
100     If Right$(FilePath, 1) <> "\" Then
102         FilePath = FilePath & "\"
        End If
         
        Dim strDirectoryName As String
    
        ' Scan Subdirs First
104     If Recursive Then
106         lngFind = FindFirstFile(FilePath & "*.*", typFindData) 'api call
108         lngMatch = 99
110         Do While lngFind > 0 And lngMatch > 0
112             If (typFindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) > 0 Then 'if this is a directory
114                 strDirectoryName = TrimToNull(typFindData.cFileName)
116                 If strDirectoryName <> "." And strDirectoryName <> ".." Then
118                     SearchFiles FilePath & strDirectoryName, FileName, Recursive 'recursive call to self
                    End If
                End If
120             lngMatch = FindNextFile(lngFind, typFindData) 'api call
            Loop

122         FindClose lngFind 'api call
        End If
    
124     lngFind = FindFirstFile(FilePath & FileName, typFindData) 'api call
126     lngMatch = 99
128     Do While lngFind > 0 And lngMatch > 0
130         If Not (typFindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) > 0 Then 'if not directory
132             Set objFile = m_Files.Add(FilePath & TrimToNull(typFindData.cFileName))

134             With objFile
136                 .FileName = TrimToNull(typFindData.cFileName)
138                 .FileAttributes = typFindData.dwFileAttributes
140                 .Created = FileTimeToDate(typFindData.ftCreationTime)
142                 .Accessed = FileTimeToDate(typFindData.ftLastAccessTime)
144                 .Modified = FileTimeToDate(typFindData.ftLastWriteTime)
146                 .FileSize = FileSizeToDouble(typFindData.nFileSizeHigh, typFindData.nFileSizeLow)
148                 .FilePath = FilePath
                End With

150             Set objFile = Nothing
            End If
152         lngMatch = FindNextFile(lngFind, typFindData) 'api call
        Loop

154     FindClose lngFind 'api call

        '<EhFooter>
ExitPoint:
        ModCallStack.StackRemove 'this method is complete, remove from call stack
    
        Exit Sub
SearchFiles_Error:

        ModCallStack.RuntimeError "Search", "SearchFiles", Err, Erl, strErrorDescription 'logging error
        Err.Clear
        Resume ExitPoint
        '</EhFooter>
End Sub

Private Sub Class_Initialize()
    Set m_Files = New ClsFiles
End Sub

Private Sub Class_Terminate()
    Set m_Files = Nothing
End Sub

Private Function TrimToNull(ByVal StringToTrim As String) As String
        If InStrB(StringToTrim, vbNullChar) > 0 Then
100         TrimToNull = Left$(StringToTrim, InStr(StringToTrim, vbNullChar) - 1)
        Else
102         TrimToNull = StringToTrim
        End If
End Function

Private Function FileTimeToDate(typFiletime As FILETIME) As Date
        '<EhHeader>
        On Error GoTo FileTimeToDate_Error
        Dim strErrorDescription As String 'string to set extended error information
        'adding to debug call stack
        ModCallStack.StackAdd "Search", "FileTimeToDate", "", ""
        '</EhHeader>
    
        Dim typSystemTime As SYSTEMTIME
    
100     FileTimeToSystemTime typFiletime, typSystemTime 'api call
    
102     FileTimeToDate = CDate(Format(typSystemTime.wMonth, "00") & "/" _
           & Format$(typSystemTime.wDay, "00") & "/" _
           & Format$(typSystemTime.wYear, "0000") & " " _
           & Format$(typSystemTime.wHour, "00") & ":" _
           & Format$(typSystemTime.wMinute, "00") & ":" _
           & Format$(typSystemTime.wSecond, "00"))
    
        '<EhFooter>
ExitPoint:
        ModCallStack.StackRemove 'this method is complete, remove from call stack
    
        Exit Function
FileTimeToDate_Error:

        ModCallStack.RuntimeError "Search", "FileTimeToDate", Err, Erl, strErrorDescription 'logging error
        Err.Clear
        Resume ExitPoint
        '</EhFooter>
End Function

Private Function FileSizeToDouble(FileSizeHigh As Long, _
                                  FileSizeLow As Long) As Double
        '<EhHeader>
        On Error GoTo FileSizeToDouble_Error
        Dim strErrorDescription As String 'string to set extended error information
        'adding to debug call stack
        ModCallStack.StackAdd "Search", "FileSizeToDouble", "", ""
        '</EhHeader>
    
100     FileSizeToDouble = CDbl((FileSizeHigh * 2147483647) + FileSizeLow)
    
        '<EhFooter>
ExitPoint:
        ModCallStack.StackRemove 'this method is complete, remove from call stack
    
        Exit Function
FileSizeToDouble_Error:

        ModCallStack.RuntimeError "Search", "FileSizeToDouble", Err, Erl, strErrorDescription 'logging error
        Err.Clear
        Resume ExitPoint
        '</EhFooter>
End Function

