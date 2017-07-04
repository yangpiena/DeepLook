Attribute VB_Name = "ModFileSearchHandler"
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:11:22
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：ModFileSearchHandler
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:11:22
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V4.12.0
'*************************************************************************

Option Explicit

'-----------------------------------------------------------------------------------------------
Private BatFiles() As String
Private BatFiles_Elements As Long

Private NodeNum As Long, AllowAdd As Boolean
'-----------------------------------------------------------------------------------------------

Public Sub LoadFiles()
    Dim I As Long, Z As Long

    On Local Error Resume Next

    BatFiles_Elements = 0

    With FrmResults.TreeView.Nodes
        For I = 1 To .Count
            ' Adds the Declared DLLs to the Copy Report
            AllowAdd = True

            NodeNum = InStr(1, .Item(I).Key, "DECDLLS_") ' Is a DecDLL
            If NodeNum <> 0 Then
                If .Item(I).Image = "DLL" Then ' Prevent System (non-copy) DLLs from being showed
                    If InStrRev(.Item(I).Key, "_") = NodeNum + 7 Then ' Only allow the root item node (no DLL information nodes)
                        For Z = 1 To FrmCopyReport.tvwItemsTV.Nodes.Count
                            If UCase$(FrmCopyReport.tvwItemsTV.Nodes(Z).Text) = UCase$(Mid$(.Item(I).Key, NodeNum + 8)) Then AllowAdd = False
                        Next

                        If AllowAdd = True And .Item(I).Image <> "CreateObject" Then
                            FrmCopyReport.tvwItemsTV.Nodes.Add , , .Item(I).Key, Mid$(.Item(I).Key, NodeNum + 8), "DLL"

                            BatFiles_Elements = BatFiles_Elements + 1
                            If UBound(BatFiles) < BatFiles_Elements Then
                                ReDim BatFiles(BatFiles_Elements + 5) As String
                            End If

                            BatFiles(BatFiles_Elements) = Mid$(.Item(I).Key, NodeNum + 8)
                        End If
                    End If
                Else
                    If InStrRev(.Item(I).Key, "_") = NodeNum + 7 Then ' Only allow the root item node (no DLL information nodes)
                        For Z = 1 To FrmCopyReport.tvwNonCopyItemsTV.Nodes.Count
                            If UCase$(FrmCopyReport.tvwNonCopyItemsTV.Nodes(Z).Text) = UCase$(Mid$(.Item(I).Key, NodeNum + 8)) Then AllowAdd = False
                        Next

                        If AllowAdd = True And .Item(I).Image <> "CreateObject" Then
                            FrmCopyReport.tvwNonCopyItemsTV.Nodes.Add , , .Item(I).Key, Mid$(.Item(I).Key, NodeNum + 8), "SysDLL"
                        End If
                    End If
                End If
            End If

            ' Adds the Reference DLLs to the Copy Report
            NodeNum = InStr(1, .Item(I).Key, "REFCOM_REFERENCE_") ' Is a DLL
            If NodeNum <> 0 Then
                If .Item(I).Image = "DLL" Then ' Prevent System (non-copy) DLLs from being showed
                    If InStrRev(.Item(I).Key, "_") = NodeNum + 16 Then
                        For Z = 1 To FrmCopyReport.tvwItemsTV.Nodes.Count
                            If UCase$(FrmCopyReport.tvwItemsTV.Nodes(Z).Text) = UCase$(Mid$(.Item(I).Key, NodeNum + 17)) Then AllowAdd = False
                        Next

                        If LCase$(Right$(FrmCopyReport.tvwItemsTV.Nodes(Z).Text, 4)) = ".vbp" Then AllowAdd = False

                        If AllowAdd = True And .Item(I).Image <> "CreateObject" Then
                            FrmCopyReport.tvwItemsTV.Nodes.Add , , .Item(I).Key, Mid$(.Item(I).Key, NodeNum + 17), "DLL"

                            BatFiles_Elements = BatFiles_Elements + 1
                            If UBound(BatFiles) < BatFiles_Elements Then
                                ReDim BatFiles(BatFiles_Elements + 5) As String
                            End If

                            BatFiles(BatFiles_Elements) = Mid$(.Item(I).Key, NodeNum + 17)
                        End If
                    End If
                Else
                    For Z = 1 To FrmCopyReport.tvwNonCopyItemsTV.Nodes.Count
                        If UCase$(FrmCopyReport.tvwNonCopyItemsTV.Nodes(Z).Text) = UCase$(Mid$(.Item(I).Key, NodeNum + 17)) Then AllowAdd = False
                    Next

                    If AllowAdd = True And InStrRev(.Item(I).Key, "_") = NodeNum + 16 And .Item(I).Image <> "CreateObject" Then
                        FrmCopyReport.tvwNonCopyItemsTV.Nodes.Add , , .Item(I).Key, Mid$(.Item(I).Key, NodeNum + 17), "SysDLL"
                    End If
                End If
            End If

            ' Adds Components to the Copy Report
            NodeNum = InStr(1, .Item(I).Key, "REFCOM_COMPONENT_") ' Is a Component
            If NodeNum <> 0 Then
                If InStrRev(.Item(I).Key, "_") = NodeNum + 16 Then
                    For Z = 1 To FrmCopyReport.tvwItemsTV.Nodes.Count
                        If UCase$(FrmCopyReport.tvwItemsTV.Nodes(Z).Text) = UCase$(Mid$(.Item(I).Key, NodeNum + 18)) Then AllowAdd = False
                    Next

                    If InStrB(1, .Item(I).Key, ".vbp", vbTextCompare) Then
                        AllowAdd = False
                        FrmCopyReport.tvwManualCopyTV.Nodes.Add , , .Item(I).Key, Trim$(Mid$(.Item(I).Key, NodeNum + 17)), "Project"
                    End If

                    If AllowAdd = True Then
                        FrmCopyReport.tvwItemsTV.Nodes.Add , , .Item(I).Key, Trim$(Mid$(.Item(I).Key, NodeNum + 17)), "Component"

                        BatFiles_Elements = BatFiles_Elements + 1
                        If UBound(BatFiles) < BatFiles_Elements Then
                            ReDim BatFiles(BatFiles_Elements + 5) As String
                        End If

                        BatFiles(BatFiles_Elements) = Mid$(.Item(I).Key, NodeNum + 18)
                    End If
                End If
            End If

            ' Add CreateObject statements to Copy Report
            If .Item(I).Image = "CreateObject" Then
                FrmCopyReport.tvwManualCopyTV.Nodes.Add , , .Item(I).Key, .Item(I).Text & " (CreateObject Statement)", "CreateObject"
            End If

            DoEvents
        Next
    End With

    If UsesGroup = True Then ' Group file, fonts need to be found in the treeview
        On Local Error Resume Next ' Dont show errors if font already added to copy report

        With FrmResults.TreeView.Nodes
            For I = 1 To .Count
                If .Item(I).Image = "Font" And .Item(I).Text <> "Used Fonts" Then
                    FrmCopyReport.tvwManualCopyTV.Nodes.Add , , .Item(I).Text, .Item(I).Text & " (Font)", "Font"
                End If
            Next
        End With

        On Local Error GoTo 0 ' Resume no error trapping
    Else ' Single project file, fonts are already in a listbox on the results form
        For I = 1 To UBound(UsedFonts) ' MUST be 1 - first item (index 0) in array is always 0 due to design
            FrmCopyReport.tvwManualCopyTV.Nodes.Add , , UsedFonts(I), UsedFonts(I) & " (Font)", "Font"
        Next
    End If
End Sub

Public Sub CopyDLLOCX()
    Dim I As Long, Z As Long

    FrmResults.sbrStatus.Caption = "STAT>拷贝注释文件..."

    With FrmCopyReport.tvwItemsTV.Nodes
        For I = 1 To .Count
            If .Item(I).Checked = False Then
                .Item(I).Image = "Skipped"

                For Z = 1 To BatFiles_Elements
                    If BatFiles(Z) = .Item(I).Text Then
                        BatFiles(Z) = BatFiles(BatFiles_Elements)
                        BatFiles_Elements = BatFiles_Elements - 1
                    End If
                Next
            Else
                Findfile Mid$(.Item(I).Key, NodeNum + 8)
            End If

            FrmCopyReport.pgbPercentBar.Value = (100 / .Count) * I
            DoEvents
        Next
    End With

    If Not FileExists(App.Path & "\FileRegister.exe") Then
        ModFileRegisterBatCreator.CreateBatHeader GetRootDirectory(ProjectPath) & "Res\FileRegister.bat"

        For I = 1 To BatFiles_Elements
            ModFileRegisterBatCreator.AddBatRegAndCopyFile BatFiles(I), I, BatFiles_Elements
        Next

        ModFileRegisterBatCreator.AddBatFooter GetRootDirectory(ProjectPath) & "Res\FileRegister.bat"
    Else
        FileCopy App.Path & "\FileRegister.exe", GetRootDirectory(ProjectPath) & "Res\FileRegister.exe"
    End If

    FrmCopyReport.Caption = "复制所需文件的报告 - 完毕."
    FrmCopyReport.btnCloseButton.Enabled = True
    FrmResults.sbrStatus.Caption = "STAT>文件复制完毕. 文件已复制到 " & GetRootDirectory(ProjectPath) & "Res\" & "."
End Sub

Private Sub Findfile(FileName As String)
    Dim FileSearch As ClsSearch, I As Long

    Set FileSearch = New ClsSearch

    If InStrB(1, FileName, ".") = 0 Then Exit Sub

    FileName = Trim$(Mid$(FileName, InStrRev(FileName, "_") + 1))

    DoEvents

    FrmResults.sbrStatus.Caption = "STAT>Commencing File Copy: " & FileName

    For I = 1 To FrmCopyReport.tvwItemsTV.Nodes.Count
        If FrmCopyReport.tvwItemsTV.Nodes(I).Text = FileName Then
            FrmCopyReport.tvwItemsTV.Nodes(I).Image = "CurrentCopy"
            Exit For
        End If
    Next

    DoEvents

    With FileSearch
        .SearchFiles Environ("windir"), FileName, True

        If .Files.Count = 0 Then AltFindFile FileName: Exit Sub

        On Error Resume Next
        Err.Clear
        FileCopy .Files.Item(1).FileNameFull, GetRootDirectory(ProjectPath) & "Res\" & .Files.Item(1).FileName
        On Error GoTo 0

        For I = 1 To FrmCopyReport.tvwItemsTV.Nodes.Count
            If FrmCopyReport.tvwItemsTV.Nodes(I).Text = FileName Then
                If Err.Number = 0 Then
                    FrmCopyReport.tvwItemsTV.Nodes(I).Image = "Done"
                Else
                    FrmCopyReport.tvwItemsTV.Nodes(I).Image = "Error"
                End If
                Exit For
            End If
        Next
    End With
End Sub

Private Sub AltFindFile(FileName As String)
    Dim FileSearch As ClsSearch, Task As Long, I As Long

    Set FileSearch = New ClsSearch

    With FileSearch
        .SearchFiles GetRootDirectory(ProjectPath), FileName, True

        If .Files.Count = 0 Then
            Task = MsgBoxEx("Cannot find file """ & FileName & """ for copying. Would you like search drive C for it?", vbYesNo, "File Copy Error", , , , , PicError)
            DoEvents

            If Task = vbYes Then
                FrmResults.sbrStatus.Caption = "STAT>Commencing File Copy: " & FileName & " [Searching C:\]"
                SearchCForFile FileName
            Else
                For I = 1 To FrmCopyReport.tvwItemsTV.Nodes.Count
                    If FrmCopyReport.tvwItemsTV.Nodes(I).Text = FileName Then
                        FrmCopyReport.tvwItemsTV.Nodes(I).Image = "Error"
                        Exit For
                    End If
                Next
                Exit Sub
            End If
        End If

        On Error Resume Next
        Err.Clear
        FileCopy .Files.Item(1).FileNameFull, GetRootDirectory(ProjectPath) & "Res\" & .Files.Item(1).FileName
        On Error GoTo 0

        For I = 1 To FrmCopyReport.tvwItemsTV.Nodes.Count
            If FrmCopyReport.tvwItemsTV.Nodes(I).Text = FileName Then
                If Err.Number = 0 Then
                    FrmCopyReport.tvwItemsTV.Nodes(I).Image = "Done"
                Else
                    FrmCopyReport.tvwItemsTV.Nodes(I).Image = "Error"
                End If
                Exit For
            End If
        Next
    End With
End Sub

Private Sub SearchCForFile(FileName As String)
    Dim FileSearch As ClsSearch, I As Long

    Set FileSearch = New ClsSearch

    With FileSearch
        .SearchFiles "C:\", FileName, True

        If .Files.Count = 0 Then
            MsgBoxEx "Cannot find file """ & FileName & """ for copying. Please manually copy this file.", vbOKOnly, "File Copy Error", , , , , PicError
            For I = 1 To FrmCopyReport.tvwItemsTV.Nodes.Count
                If FrmCopyReport.tvwItemsTV.Nodes(I).Text = FileName Then
                    FrmCopyReport.tvwItemsTV.Nodes(I).Image = "Error"
                    Exit For
                End If
            Next

            Exit Sub
        End If

        On Error Resume Next
        Err.Clear
        FileCopy .Files.Item(1).FileNameFull, GetRootDirectory(ProjectPath) & "Res\" & .Files.Item(1).FileName
        On Error GoTo 0

        For I = 1 To FrmCopyReport.tvwItemsTV.Nodes.Count
            If FrmCopyReport.tvwItemsTV.Nodes(I).Text = FileName Then
                If Err.Number = 0 Then
                    FrmCopyReport.tvwItemsTV.Nodes(I).Image = "Done"
                Else
                    FrmCopyReport.tvwItemsTV.Nodes(I).Image = "Error"
                End If

                Exit For
            End If
        Next
    End With
End Sub
