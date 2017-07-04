Attribute VB_Name = "ModVariableHandler"
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:11:49
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：ModVariableHandler
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:11:49
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V4.12.0
'*************************************************************************

'-----------------------------------------------------------------------------------------------
'                                 VB6 UNUSED VARIABLE SCANNER SYNOPSIS
'-----------------------------------------------------------------------------------------------
'   This module contains a simple and fast unused variable scanner. It is designed to be as
'   accurate as possible without sacrificing speed. It does not scan files for unused SPF
'   parameters or constants.
'
'   Many programmers like to preserve the case of their enum and type members by encasing bogus
'   declarations of the same names inside "#If False Then" statements. DeepLook recognises these
'   and ignores all variables declared inside these until it encounters a "#End If" statement.
'
'   The VB unused variable scanning engine contained in this module is (C) Dean Camera.
'-----------------------------------------------------------------------------------------------

Option Explicit

'-----------------------------------------------------------------------------------------------
Private CurrFilename As String
Private CurrProjName As String

Public GlobalVars() As String
Public GlobalVarsLoc() As String
Public GlobalVars_Elements As Long

Private LocalVars() As String
Private LocalVarsLoc() As String
Private LocalVars_Elements As Long

Private SPFVars() As String
Private SPFVarsLoc() As String
Private SPFVars_Elements As Long

Private ProjRoot As String
Private Project1Percent As Single

Private InSPF As Boolean
Private SPFName As String

Private Lop As Integer
Private InCaseProcIf As Boolean

Private PreviousLine As String
Private IsLineCont As Boolean
'-----------------------------------------------------------------------------------------------

Public Sub ClearGlobals()
    ReDim GlobalVars(20) As String
    ReDim GlobalVarsLoc(20) As String

    GlobalVars_Elements = 0 ' Clear the number of variables in the array variable
End Sub

Public Sub ClearLocals()
    ReDim LocalVarsLoc(20) As String
    ReDim LocalVars(20) As String

    ReDim SPFVarsLoc(20) As String
    ReDim SPFVars(20) As String

    LocalVars_Elements = 0 ' Clear the number of variables in the array variable
    SPFVars_Elements = 0 ' Clear the number of variables in the array variable
End Sub

Public Sub AnalyseVBProjectForVars(ProjectFileName As String)
    Dim linedata As String, LoopVar As Long, DonePrj As Long, ProjFileNum As Integer

    linedata = UCase$(Right$(ProjectFileName, 3))
    If linedata <> "VBG" And linedata <> "VBP" Then
        CurrProjName = "(Single File)"
        ScanFile ProjectFileName, False
        Exit Sub
    End If

    If IsExit Then Exit Sub ' If quitting, don't process the project for unused variables

    ProjectFileName = FixRelPath(ProjectFileName) ' Make path literal if relative

    If Not FileExists(FilesRootDirectory & ProjectFileName) And Not FileExists(ProjectFileName) Then
        Exit Sub ' Just exit if project file not found, warning would already have been given by code scanner
    End If

    CurrProjName = ExtractFileName(ProjectFileName)
    ProjRoot = Mid$(ProjectFileName, 1, InStrRev(ProjectFileName, "\"))
    ProjFileNum = FreeFile

    If FileExists(FilesRootDirectory & ProjectFileName) Then
        Open FilesRootDirectory & ProjectFileName For Input As #ProjFileNum
    Else
        Open ProjectFileName For Input As #ProjFileNum
    End If

    With FrmSelProject
        .pgbAPB.Value = 0
        .pgbAPB.Color = 8421631
        .lblScanPhase.Caption = "Unused Variable Scan Phase"
        .lblScanPhase.ForeColor = 8421631
        .imgCurrScanObjType.Picture = FrmResults.ilstImages.ListImages(29).Picture
    End With

    DoEvents

    Project1Percent = (100 / LOF(ProjFileNum))

    Do While Not EOF(ProjFileNum)
        Line Input #ProjFileNum, linedata ' Get the line data from the project file

        LookAtPRJLine linedata
        FrmSelProject.pgbAPB.Value = Project1Percent * DonePrj
        DonePrj = DonePrj + Len(linedata)

        If GetInputState Then DoEvents ' Process system events if there is messages in the keyboard/mouse buffers
    Loop

    Close #ProjFileNum

    With FrmResults.lstVarList
        For LoopVar = 1 To GlobalVars_Elements ' Add unused global variables to the list
            .ListItems.Add , , CurrProjName
            .ListItems(.ListItems.Count).ListSubItems.Add , , GlobalVarsLoc(LoopVar)
            .ListItems(.ListItems.Count).ListSubItems.Add , , "Global"
            .ListItems(.ListItems.Count).ListSubItems.Add , , GlobalVars(LoopVar)

            .ListItems(.ListItems.Count).ListSubItems(2).ForeColor = RGB(140, 10, 140)
        Next
    End With
    
    GlobalVars_Elements = 0
End Sub

Public Sub ScanFile(linedata As String, Optional IsProject As Boolean) ' No statistics are being generated, so the Header dosn't need to be ignored
    Dim TempVar As Long, FileNum As Integer

    If InStrB(1, linedata, ";") Then linedata = Mid$(linedata, InStr(1, linedata, ";") + 1)

    On Local Error Resume Next
    linedata = Trim$(linedata)
    InSPF = False ' Reset flag; used to tell if the scanning engine is inside a sub/function/property
    CurrFilename = vbNullString

    FileNum = FreeFile

    If Not FileExists(FilesRootDirectory & ProjRoot & linedata) Then
        If Not FileExists(ProjRoot & linedata) Then
            If Not FileExists(linedata) Then
                FrmResults.lstVarList.ListItems.Add , , CurrProjName
                FrmResults.lstVarList.ListItems(FrmResults.lstVarList.ListItems.Count).ForeColor = vbRed

                With FrmResults.lstVarList.ListItems(FrmResults.lstVarList.ListItems.Count)
                    .ListSubItems.Add , , linedata
                    .ListSubItems(1).ForeColor = vbRed
                    .ListSubItems.Add , , "ERROR"
                    .ListSubItems(2).ForeColor = vbRed
                    .ListSubItems.Add , , "The requested file could not be found."
                    .ListSubItems(3).ForeColor = vbRed
                End With

                Exit Sub
            Else
                Open linedata For Input As #FileNum
            End If
        Else
            Open ProjRoot & linedata For Input As #FileNum
        End If
    Else
        Open FilesRootDirectory & ProjRoot & linedata For Input As #FileNum
    End If

    Do While Not EOF(FileNum)
        If IsExit Then Exit Do ' If quitting, don't process any more of the file

        Line Input #FileNum, linedata ' Get the file line data

        If Right$(linedata, 1) = "_" Then
            PreviousLine = PreviousLine & Left$(linedata, Len(linedata) - 1)
            IsLineCont = True
        Else
            If IsLineCont Then
                linedata = PreviousLine & linedata
                PreviousLine = vbNullString
                IsLineCont = False
            End If

            If LenB(CurrFilename) = 0 Then
                If Left$(linedata, 20) = "Attribute VB_Name = " Then CurrFilename = Mid$(linedata, 21) ' Get the filename if nessesary
            ElseIf Not IsCommentLine(linedata) Then
                If InStrB(1, linedata, "'") Then
                    TempVar = IsHybridLine(linedata)

                    If TempVar Then linedata = Left$(linedata, TempVar) ' Remove comments if nessesary
                End If

                linedata = Trim$(linedata)

                If CheckIfIsEndStartSPF(linedata) = 0 Then ' If line is not the beginning or end of a SPF, scan it for variables
                    CheckIfVarIsUsed linedata

                    If Left$(linedata, 14) = "#If False Then" Then ' These statements encase bogus declarations that preserve the case of type/enum members
                        InCaseProcIf = True
                    ElseIf linedata = "#End If" Then ' End of case protection
                        InCaseProcIf = False
                    ElseIf Not InCaseProcIf Then  ' If not inside a protection IF, check for variable declarations
                        CheckIfIsDeclare linedata
                    End If
                End If
            End If
        End If
    Loop

    Close #FileNum

    For TempVar = 1 To LocalVars_Elements ' Add all the unused local variables to the list
        If LenB(Trim$(LocalVarsLoc(TempVar))) = 0 Or Len(Trim$(LocalVars(TempVar))) = 0 Then
            ' Don't use a <> operator, only works like this for some reason
        Else
            With FrmResults.lstVarList
                .ListItems.Add , , CurrProjName
                .ListItems(.ListItems.Count).ListSubItems.Add , , LocalVarsLoc(TempVar)
                .ListItems(.ListItems.Count).ListSubItems.Add , , "Local"
                .ListItems(.ListItems.Count).ListSubItems.Add , , LocalVars(TempVar)

                .ListItems(.ListItems.Count).ListSubItems(2).ForeColor = RGB(10, 10, 220)
            End With
        End If
    Next
    
    LocalVars_Elements = 0
End Sub

Private Sub LookAtPRJLine(linedata As String)
    If Left$(linedata, 5) = "Form=" Then
        ClearLocals
        ScanFile Mid$(linedata, 6), True
    ElseIf Left$(linedata, 7) = "Module=" Then
        ClearLocals
        ScanFile Mid$(linedata, 8), True
    ElseIf Left$(linedata, 6) = "Class=" Then
        ClearLocals
        ScanFile Mid$(linedata, 7), True
    ElseIf Left$(linedata, 12) = "UserControl=" Then
        ClearLocals
        ScanFile Mid$(linedata, 13), True
    ElseIf Left$(linedata, 13) = "PropertyPage=" Then
        ClearLocals
        ScanFile Mid$(linedata, 14), True
    ElseIf Left$(linedata, 13) = "UserDocument=" Then
        ClearLocals
        ScanFile Mid$(linedata, 14), True
    ElseIf Left$(linedata, 9) = "Designer=" Then
        ClearLocals
        ScanFile Mid$(linedata, 10), True
    End If
End Sub

Private Function CheckVar(VType As Integer, I As Long, linedata As String) As Boolean
    Dim SearchFor As String, TempByte As String * 1, TempInt As Long, StartPos As Long

    StartPos = 1

    If VType = 1 Then
        SearchFor = GlobalVars(I)
    ElseIf VType = 2 Then
        SearchFor = LocalVars(I)
    ElseIf VType = 3 Then
        SearchFor = SPFVars(I)
    End If

Recheck:
    
    TempInt = InStr(StartPos, linedata, SearchFor)
    If TempInt = 0 Then Exit Function

    If TempInt > 1 Then
        TempByte = Mid$(linedata, TempInt - 1, 1)
        If InStrB(1, " (.#", TempByte) Then GoTo CheckEnd
    Else
        GoTo CheckEnd
    End If

    StartPos = StartPos + 1
    
    GoTo Recheck
CheckEnd:

    If Len(linedata) > (TempInt + Len(SearchFor)) Then
        TempByte = Mid$(linedata, TempInt + Len(SearchFor), 1)
        If InStrB(1, ", )(.&#$%!", TempByte) Then GoTo IncUsed
    Else
        GoTo IncUsed
    End If

    StartPos = StartPos + 1
    GoTo Recheck
IncUsed:

    CheckVar = True

    Select Case VType
        Case 1
            GlobalVars(I) = GlobalVars(GlobalVars_Elements)
            GlobalVarsLoc(I) = GlobalVarsLoc(GlobalVars_Elements)

            GlobalVars_Elements = GlobalVars_Elements - 1
        Case 2
            LocalVars(I) = LocalVars(LocalVars_Elements)
            LocalVarsLoc(I) = LocalVarsLoc(LocalVars_Elements)

            LocalVars_Elements = LocalVars_Elements - 1
        Case 3
            SPFVars(I) = SPFVars(SPFVars_Elements)
            SPFVarsLoc(I) = SPFVarsLoc(SPFVars_Elements)

            SPFVars_Elements = SPFVars_Elements - 1
    End Select
End Function

Private Sub CheckIfVarIsUsed(linedata As String)
    Dim LoopVar As Long

ScanGlobals:
    For LoopVar = 1 To GlobalVars_Elements
        If InStrB(1, linedata, GlobalVars(LoopVar)) Then
            If Left$(linedata, 7) <> "Public " And Left$(linedata, 7) <> "Global " Then ' Stops it from interpreting the variable's own declaration as it's use
                If CheckVar(1, LoopVar, linedata) Then GoTo ScanGlobals ' If variable used, rescan line again to see if other variables are also used
            End If
        End If
    Next

ScanLocals:
    For LoopVar = 1 To LocalVars_Elements
        If InStrB(1, linedata, LocalVars(LoopVar)) Then
            If CheckVar(2, LoopVar, linedata) Then GoTo ScanLocals ' If variable used, rescan line again to see if other variables are also used
        End If
    Next

ScanSPFs:
    For LoopVar = 1 To SPFVars_Elements
        If InStrB(1, linedata, SPFVars(LoopVar)) Then
            If CheckVar(3, LoopVar, linedata) Then GoTo ScanSPFs ' If variable used, rescan line again to see if other variables are also used
        End If
    Next
End Sub

Private Sub CheckIfIsDeclare(linedata As String)
    If Left$(linedata, 4) = "Dim " Then
        Lop = 4
    ElseIf Left$(linedata, 8) = "Private " Then
        Lop = 8
    ElseIf Left$(linedata, 7) = "Static " Then
        Lop = 7
    ElseIf Left$(linedata, 11) = "WithEvents " Then
        Lop = 12
    ElseIf Left$(linedata, 6) = "Const " Then
        Exit Sub
    Else
        Exit Sub
    End If

    If InStrB(1, linedata, "Declare Function ") Then
        Exit Sub
    ElseIf InStrB(1, linedata, " Const ") Then
        Exit Sub
    ElseIf InStrB(1, linedata, " Type ") Then
        Exit Sub
    ElseIf InStrB(1, linedata, " Enum ") Then
        Exit Sub
    ElseIf InStrB(1, linedata, " Event ") Then
        Exit Sub
    ElseIf InStrB(1, linedata, "Declare Sub ") Then
        Exit Sub
    ElseIf InStrB(1, linedata, " Property ") Then
        Exit Sub
    End If

    If InStrB(1, linedata, " WithEvents ") Then linedata = Replace$(linedata, " WithEvents ", vbNullString, 1, 1) ' Remove "WithEvents" if present

    If InSPF Then
        AddSPFVar linedata
    Else
        AddLocalVar linedata
    End If
End Sub

Private Function FixQuotes(Data As String) As String
    If Left$(Data, 1) = """" Then
        Data = Mid$(Data, 2)
        If Right$(Data, 1) = """" Then Data = Left$(Data, Len(Data) - 1)
    End If

    FixQuotes = Data
End Function

Private Function CheckIfIsEndStartSPF(linedata As String) As Boolean
    Dim LoopVar As Long, SubN As String, FuncN As String, PropN As String

    If Left$(linedata, 4) = "End " Then
        linedata = Trim$(Mid$(linedata, 5))

        If linedata = "Sub" Or linedata = "Function" Or linedata = "Property" Then
            For LoopVar = 1 To SPFVars_Elements
                If LenB(Trim$(SPFVarsLoc(LoopVar))) = 0 Or Len(Trim$(SPFVars(LoopVar))) = 0 Then
                    ' Don't use a <> operator, only works like this for some reason
                Else
                    With FrmResults.lstVarList
                        .ListItems.Add , , CurrProjName
                        .ListItems(.ListItems.Count).ListSubItems.Add , , SPFVarsLoc(LoopVar)
                        .ListItems(.ListItems.Count).ListSubItems.Add , , SPFName
                        .ListItems(.ListItems.Count).ListSubItems.Add , , SPFVars(LoopVar)

                        .ListItems(.ListItems.Count).ListSubItems(2).ForeColor = RGB(230, 110, 110)

                        If InStrB(1, SPFVars(LoopVar), " = ") Then .ListItems(.ListItems.Count).ListSubItems(3).ForeColor = RGB(50, 150, 50)
                    End With
                End If
            Next

            SPFVars_Elements = 0

            InSPF = False
            CheckIfIsEndStartSPF = True
        End If
    ElseIf InStrB(1, linedata, "(") Then
        SubN = CheckIsSub(linedata)
        PropN = CheckIsFunction(linedata)
        FuncN = CheckIsProperty(linedata)

        If LenB(SubN) Or LenB(PropN) Or LenB(FuncN) Then  ' New SPF starting
            If LenB(SubN) Then
                SPFName = SubN
            ElseIf LenB(PropN) Then
                SPFName = PropN
            ElseIf LenB(FuncN) Then
                SPFName = FuncN
            End If

            SPFName = Trim$(Left$(SPFName, InStr(1, SPFName, "(") - 1)) ' Remove parameters from SPF name

            InSPF = True ' Scanning inside a SPF; set flag to true
            CheckIfIsEndStartSPF = True
        End If
    End If
End Function

Private Sub AddLocalVar(linedata As String)
    Dim VarNames() As String, Var As Long, TempVar As Long

    While InStrB(1, linedata, "(") ' Remove any array information from the variable(s)
        TempVar = InStr(1, linedata, ")")
        linedata = Left$(linedata, InStr(1, linedata, "(") - 1) & Mid$(linedata, TempVar + 1)
    Wend

    If UBound(LocalVars) < LocalVars_Elements Then
        ReDim Preserve LocalVars(LocalVars_Elements + 20) As String
        ReDim Preserve LocalVarsLoc(LocalVars_Elements + 20) As String
    End If

    If InStrB(1, linedata, ",") = 0 Then
        LocalVars_Elements = LocalVars_Elements + 1
        
        If InStrB(1, linedata, " As ") Then
            LocalVars(LocalVars_Elements) = TrimJunk(Mid$(linedata, Lop, InStr(1, linedata, " As ") - Lop))
        Else
            LocalVars(LocalVars_Elements) = TrimJunk(Mid$(linedata, Lop))
        End If

        LocalVarsLoc(LocalVars_Elements) = FixQuotes(CurrFilename)
    Else
        linedata = Trim$(Mid$(linedata, Lop))
        If InStrB(1, linedata, ":") Then linedata = Left$(linedata, InStr(1, linedata, ":"))

        VarNames = Split(linedata, ",")
        For Var = LBound(VarNames) To UBound(VarNames)
            LocalVars_Elements = LocalVars_Elements + 1
            
            If InStrB(1, VarNames(Var), " As ") Then
                LocalVars(LocalVars_Elements) = TrimJunk(Left$(VarNames(Var), InStr(1, VarNames(Var), " As ")))
            Else
                LocalVars(LocalVars_Elements) = TrimJunk(VarNames(Var))
            End If

            LocalVarsLoc(LocalVars_Elements) = FixQuotes(CurrFilename)
        Next
    End If
End Sub

Private Sub AddSPFVar(linedata As String)
    Dim VarNames() As String, Var As Long, TempVar As Long
        
    While InStrB(1, linedata, "(") ' Remove any array information from the variable(s)
        TempVar = InStr(1, linedata, ")")
        linedata = Left$(linedata, InStr(1, linedata, "(") - 1) & Mid$(linedata, TempVar + 1)
    Wend

    If UBound(SPFVars) < SPFVars_Elements Then
        ReDim Preserve SPFVars(SPFVars_Elements + 20) As String
        ReDim Preserve SPFVarsLoc(SPFVars_Elements + 20) As String
    End If
    
    If InStrB(1, linedata, ",") = 0 Then
        SPFVars_Elements = SPFVars_Elements + 1
        
        If InStrB(1, linedata, " As ") Then
            SPFVars(SPFVars_Elements) = TrimJunk(Mid$(linedata, Lop, InStr(1, linedata, " As ") - Lop))
        Else
            SPFVars(SPFVars_Elements) = TrimJunk(Mid$(linedata, Lop))
        End If

        SPFVarsLoc(SPFVars_Elements) = FixQuotes(CurrFilename)
    Else
        linedata = Trim$(Mid$(linedata, Lop))
        If InStrB(1, linedata, ":") Then linedata = Left$(linedata, InStr(1, linedata, ":"))

        VarNames = Split(linedata, ",")
        For Var = LBound(VarNames) To UBound(VarNames)
            SPFVars_Elements = SPFVars_Elements + 1
            
            If InStrB(1, VarNames(Var), " As ") Then
                SPFVars(SPFVars_Elements) = TrimJunk(Left$(VarNames(Var), InStr(1, VarNames(Var), " As ")))
            Else
                SPFVars(SPFVars_Elements) = TrimJunk(VarNames(Var))
            End If

            SPFVarsLoc(SPFVars_Elements) = FixQuotes(CurrFilename)
        Next
    End If
End Sub

Private Function CheckIsSub(linedata As String) As String ' Returns Sub name if the line is a Sub
    If InStrB(1, linedata, "Sub ", vbBinaryCompare) Then
        If InStrB(1, linedata, "(") Then
            If InStrB(1, linedata, "Declare ") = 0 Then
                If Left$(linedata, 12) = "Private Sub " Then
                    CheckIsSub = Mid$(linedata, 12)
                ElseIf Left$(linedata, 11) = "Public Sub " Then
                    CheckIsSub = Mid$(linedata, 11)
                ElseIf Left$(linedata, 4) = "Sub " Then
                    CheckIsSub = Mid$(linedata, 4)
                ElseIf Left$(linedata, 11) = "Friend Sub " Then
                    CheckIsSub = Mid$(linedata, 11)
                ElseIf Left$(linedata, 11) = "Static Sub " Then
                    CheckIsSub = Mid$(linedata, 11)
                ElseIf Left$(linedata, 19) = "Public Static Sub " Then
                    CheckIsSub = Mid$(linedata, 19)
                ElseIf Left$(linedata, 18) = "Private Static Sub " Then
                    CheckIsSub = Mid$(linedata, 18)
                End If
            End If
        End If
    End If
End Function

Private Function CheckIsFunction(linedata As String) As String ' Returns Function name if the line is a Function
    If InStrB(1, linedata, "Function ", vbBinaryCompare) Then
        If InStrB(1, linedata, "(") Then
            If InStrB(1, linedata, "Declare ", vbBinaryCompare) = 0 Then
                If Left$(linedata, 8) = "Private " Then
                    If Mid$(linedata, 9, 7) = "Static " Then
                        CheckIsFunction = Mid$(linedata, 24)
                    Else
                        CheckIsFunction = Mid$(linedata, 17)
                    End If
                ElseIf Left$(linedata, 7) = "Public " Then
                    If Mid$(linedata, 9, 7) = "Static " Then
                        CheckIsFunction = Mid$(linedata, 23)
                    Else
                        CheckIsFunction = Mid$(linedata, 16)
                    End If
                ElseIf Left$(linedata, 9) = "Function " Then
                    CheckIsFunction = Mid$(linedata, 9)
                ElseIf Left$(linedata, 7) = "Friend " Then
                    CheckIsFunction = Mid$(linedata, 17)
                ElseIf Left$(linedata, 7) = "Static " Then
                    CheckIsFunction = Mid$(linedata, 17)
                End If
            End If
        End If
    End If
End Function

Private Function CheckIsProperty(linedata As String) As String ' Returns Property name if the line is a Property
    If InStrB(1, linedata, "Property ", vbBinaryCompare) Then
        If InStrB(1, linedata, "(") Then
            If Left$(linedata, 17) = "Private Property " Then
                CheckIsProperty = Mid$(linedata, 21)
            ElseIf Left$(linedata, 16) = "Public Property " Then
                CheckIsProperty = Mid$(linedata, 20)
            ElseIf Left$(linedata, 9) = "Property " Then
                CheckIsProperty = Mid$(linedata, 13)
            End If
        End If
    End If
End Function
