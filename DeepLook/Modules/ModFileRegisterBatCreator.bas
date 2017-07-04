Attribute VB_Name = "ModFileRegisterBatCreator"
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:11:18
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：ModFileRegisterBatCreator
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:11:18
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V4.12.0
'*************************************************************************

Option Explicit

'-----------------------------------------------------------------------------------------------
Dim BATFileNum As Integer
'-----------------------------------------------------------------------------------------------

Public Sub CreateBatHeader(FileName As String)
    BATFileNum = FreeFile
    Open FileName For Output As #BATFileNum

    Print #BATFileNum, "@echo off" & vbCrLf & "echo                    En-Tech DeepLook Project Scanner" & vbCrLf & _
        "echo             *** Automatic file register batch script file ***" & vbCrLf & _
        "echo -------------------------------------------------------------------------------" & vbCrLf & _
        "echo You must be using WinME/98/95 and have the RegSvr32.exe in your windows folder." & vbNewLine & "echo." & _
        vbCrLf & "pause" & vbCrLf & "cls"
End Sub

Public Sub AddBatRegAndCopyFile(FileName As String, Findex As Long, Fmax As Long)
    Print #BATFileNum, "echo *** Copying File #" & Findex & " of " & Fmax & " (" & FileName & ")..."
    Print #BATFileNum, "echo." & vbCrLf & "copy """ & FileName & """, """ & "%WINDIR%\System\" & FileName & """"
    Print #BATFileNum, "echo *** Registering File #" & Findex & " of " & Fmax & " (" & FileName & ")..."
    Print #BATFileNum, "%WINDIR%\System\Regsvr32.exe ""%WINDIR%\System\" & FileName & """ /s"
    Print #BATFileNum, "wait 1" & vbCrLf & "cls"
End Sub

Public Sub AddBatFooter(FileName As String)
    Print #BATFileNum, "echo." & vbCrLf & "echo." & vbCrLf & "echo File copy/registration complete." & vbCrLf & "pause"
    
    Close #BATFileNum
End Sub
