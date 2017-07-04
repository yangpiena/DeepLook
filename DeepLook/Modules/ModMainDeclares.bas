Attribute VB_Name = "ModMainDeclares"
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:11:33
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：ModMainDeclares
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:11:33
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V4.12.0
'*************************************************************************

Option Explicit

' -----------------------------------------------------------------------------------------------
                        Public Const ShowItemKey As Boolean = False
'  Change to TRUE for DEBUG purposes - shows node internal item key instead of text in statusbar
' -----------------------------------------------------------------------------------------------

Public Declare Function GetInputState Lib "user32" () As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (hwnd As Long, wMsg As Long, wParam As Long, lParam As Any) As Long
Public Const WM_SETREDRAW = &HB
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Private Declare Sub ReleaseCapture Lib "user32" ()

' -----------------------------------------------------------------------------------------------
Public PROJECTMODE As PMode
Public UsesGroup As Boolean

Public IsScanning As Boolean
Public IsExit As Boolean

Public ShowCurrItemPic As Long
Public TotalLines As Long

Public ProjectPath As String
Public FilesRootDirectory As String
Public EXENewOrOld As String * 2

Public DefaultStatText As String

Public Enum PMode
    VB6 = 0
    NET = 1
End Enum

#If False Then ' Learnt from Roger Gilchrist's Code Fixer - preserves the case of
    Private VB6, NET ' Enum statements when typing them into the IDE
#End If
' -----------------------------------------------------------------------------------------------

Public Sub Main()

    If GetSetting("DLAddin", "Options", "EXEpath", vbNullString) = vbNullString Then
        ' This notifies the DeepLook addin of DeepLook's path if it has not been set
        SaveSetting "DLAddin", "Options", "EXEpath", IIf(Right$(App.Path, 1) = "\", App.Path, App.Path & "\") & App.EXEName
    End If

    With FrmSelProject
        If LenB(Command$) Then ' Command line specified, scan passed filename
            .cmbProjectPath.Text = Command$ ' Set the file path
            .btnGoAnalyse.Enabled = False
            .Show  ' Show the select project form
            DoEvents ' Wait for the form to show before continuing
            .GoAnalyse
            .btnGoAnalyse.Enabled = True
        Else
            .Show  ' Show the select project form
        End If
    End With
End Sub

Public Function FixRelPath(Path As String) As String
    Dim FirstPartLen As Long, LastPartStart As Long, LastPart As String

    On Local Error Resume Next
    Err.Clear

    If InStrB(1, Path, "\..") Then
        Do
            FirstPartLen = InStr(1, Path, "\..")

            LastPartStart = InStr(FirstPartLen + 1, Path, "\")
            If LastPartStart Then
                LastPart = Mid$(Path, LastPartStart)
            Else
                LastPart = vbNullString
            End If

            Path = Left$(Path, InStrRev(Path, "\", FirstPartLen - 1) - 1) & LastPart
            DoEvents

            If Err.Number Then Exit Do
        Loop While InStrB(1, Path, "\..")
    End If

    FixRelPath = Path
End Function

Public Function FileExists(Path As String) As Boolean 'Returns TRUE of the file exists. Can be used for paths or files.
    On Local Error Resume Next
    FileExists = LenB(Dir(Path))
    On Local Error GoTo 0
End Function

Public Sub AddReportHeader() ' Adds the start of the report text to the report Rich-Text control
    AddReportText "=============================================================", True
    AddReportText "===================DEEPLOOK PROJECT REPORT==================="
    AddReportText "============================================================="
    AddReportText "==========    Made with DeepLook Version: " & App.Major & "." & App.Minor & "." & App.Revision & "   =========="
    AddReportText "============================================================="
    AddReportText "= You will need the ""Courier New"" font to view this report. ="
    AddReportText "============================================================="
    AddReportText "=         REPORT MADE ON: " & Now & Space$(34 - Len(Now)) & "="
    AddReportText "============================================================="
End Sub

Public Sub AddReportFooter() ' Add the end of the Report to the report Rich-Text control
    AddReportText vbNewLine & vbNewLine & "============================================================="
    AddReportText "================ END OF DEEPLOOK SCAN REPORT ================"
    AddReportText "============================================================="
    AddReportText "===========    Made by Dean Camera, 2003-2005    ============"
    AddReportText "====    Address all emails to dean_camera@hotmail.com    ===="
    AddReportText "============================================================="
End Sub

Public Sub AddReportText(AddText As String, Optional NoAddNL As Boolean) ' Adds a new line and the inputted text to the report
    FrmReport.rtbReportText.SelStart = Len(FrmReport.rtbReportText.Text)

    If Not NoAddNL Then  ' Add a new line at the start of the string
        FrmReport.rtbReportText.SelText = vbNewLine & AddText
    Else ' Don't add new line
        FrmReport.rtbReportText.SelText = AddText
    End If
End Sub

Public Function IsSysDLL(FileName As String) As String
    Dim I As Long

    Err.Clear ' Clear the error object
    On Local Error Resume Next ' Start silent error trapping

    While Err.Number = 0 ' When end of string table reached, error will exit the do loop
        I = I + 1 ' Increment the string index variable
        If FileName = LoadResString(100 + I) Then ' Stored sys DLL identical to current filename
            IsSysDLL = "SysDLL" ' Return SysDLL
            Exit Function
        End If
    Wend

    On Local Error GoTo 0 ' Resume no error trapping

    IsSysDLL = "DLL" ' Function hasn't exited, not a system DLL
End Function

Public Sub KillProgram(Unloading As Boolean)
    Dim TempXMLFilePath As String

    If IsExit = True And Not Unloading And Not IsScanning Then
        ModAnalyseVB6.PostScanDeleteObjects ' Delete any remaining instances made by the VB6 scanning engine
        IsExit = False ' Prevents recursion when the forms are unloaded

        TempXMLFilePath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & "DeepLookXMLTemp.xml"

        If FileExists(TempXMLFilePath) Then ' Temp XML report (generated when report form opened) exists
            Kill TempXMLFilePath ' Delete the temp XML report
        End If

        Dim oFrm As Form

        For Each oFrm In VB.Forms
            Unload oFrm
            Set oFrm = Nothing
        Next
    End If
End Sub

Public Function GetRootDirectory(FileName As String) As String 'Retrieves only the path (not the filename) from a Path & Filename string.
    Dim SlashPos As Long

    SlashPos = InStrRev(FileName, "\") 'Get position of the last directory slash in the string

    If SlashPos Then
        GetRootDirectory = Left$(FileName, SlashPos) 'Trim to only the directory, using the position of the last known directory slash
        If LenB(GetRootDirectory) Then
            If Right$(GetRootDirectory, 1) <> "\" Then
                GetRootDirectory = GetRootDirectory & "\" 'If there is no directory slash at the end of the string,
            End If '                                       add it to make all returned strings uniform.
        End If
    End If
End Function

Public Function ExtractFileName(FileNameAndPath As String) As String 'Gets only the filename from a file & path string
    Dim SlashPos As String

    SlashPos = InStrRev(FileNameAndPath, "\") 'Retrieve the position of the last path slash
    ExtractFileName = FileNameAndPath

    If SlashPos Then
        ExtractFileName = Mid$(FileNameAndPath, SlashPos + 1) 'Trim path to only the filename
    End If
End Function

Public Function IsHybridLine(linedata As String) As Long ' Returns TRUE if the entered string contains both code and comments
    Dim InString As Boolean, LoopVar As Long

    If InStrB(1, linedata, "'") = 0 Then
        Exit Function ' Comments after code can only use the "'" symbol, not the REM statement. Check to see if the line contains the "'" character, otherwise skip the sub to save time.
    ElseIf InStrB(1, linedata, """") = 0 Then
        IsHybridLine = InStr(linedata, "'") - 1
        Exit Function ' If there's no quotes, the line is automatically a hybrid
    End If

    For LoopVar = 2 To Len(linedata)
        Select Case Mid$(linedata, LoopVar, 1)
            Case """"
                InString = Not InString ' Changes the InString boolean variable as DeepLook finds a quote (") symbol. This prevents the program mis-interpreting "'" symbols in strings as comments.
            Case "'"
                If Not InString Then
                    IsHybridLine = LoopVar
                    Exit Function
                End If
        End Select
    Next
End Function

Public Function IsCommentLine(linedata As String) As Boolean  ' Returns TRUE if the entered line contains ONLY a comment.
    If Left$(linedata, 1) = "'" Then
        IsCommentLine = True
    ElseIf Left$(linedata, 4) = "Rem " Then
        IsCommentLine = True
    End If
End Function

Public Function TrimJunk(ByVal Data As String) As String ' Trims surrounding whitespace, procedure parameters and variable type symbols
    Data = Trim$(Data)

    If InStrB(1, Data, "(") Then Data = Mid$(Data, 1, InStr(1, Data, "(") - 1)
    If InStrB(1, "%$#&!", Right$(Data, 1)) Then Data = Left$(Data, Len(Data) - 1)
    If Left$(Data, 1) = "#" Then Data = Mid$(Data, 2)

    TrimJunk = Data
End Function

Public Function OneIfNull(Data As Long) As Long ' Used to prevent divide by zero errors if data is 0
    OneIfNull = Data
    If Data = 0 Then OneIfNull = 1
End Function

Public Sub RemoveTabStops(Frm As Form)
    Dim I As Integer

    On Local Error Resume Next

    For I = 1 To Frm.Controls.Count ' Remove all tabstops
        Frm.Controls.TabStop = False
    Next

    On Local Error GoTo 0
End Sub

Public Function DragForm(hwnd As Long) ' Drags the hWnd form when the left mouse button clicked and moved
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Function

Public Sub BubbleSortArray(ArrayToSort() As String) ' Sorts a variable length array into alphabetical order (simple bubble sort)
    Dim Index1 As Long, Index2 As Long
    Dim UpperBound As Long, TempStr As String

    ResizeArray ArrayToSort ' Important: Remove slack before sorting
    UpperBound = UBound(ArrayToSort) - 1

    For Index1 = 0 To UpperBound - 1
        For Index2 = Index1 + 1 To UpperBound
            If StrComp(ArrayToSort(Index1), ArrayToSort(Index2), vbTextCompare) = 1 Then
                TempStr = ArrayToSort(Index1)

                ArrayToSort(Index1) = ArrayToSort(Index2)
                ArrayToSort(Index2) = TempStr
            End If
        Next
    Next
End Sub

Public Sub ResizeArray(ByRef InpArray() As String) ' Removes blank (memory slack) elements from an array
    Dim UpperBound As Long
    Dim I As Long

    UpperBound = UBound(InpArray) ' Get the upper bound of the array to resize

    For I = UpperBound To 1 Step -1 ' Slack is at the end of the array, so start from there and work backwards
        If InpArray(I) <> vbNullString Then ' End of slack reached
            ReDim Preserve InpArray(I) As String ' Array is ByRef, so the passed array is actually changed in memory
            Exit Sub ' Stop processing array
        End If
    Next
End Sub
