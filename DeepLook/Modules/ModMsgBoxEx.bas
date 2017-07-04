Attribute VB_Name = "ModMsgBoxEx"
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:11:38
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：ModMsgBoxEx
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:11:38
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V4.12.0
'*************************************************************************
' THIS MODULE WAS NOT WRITTEN BY DEAN CAMERA. I CANNOT OFFER ANY SUPPORT FOR THIS MODULE.

'****************************************************************************************
'Module:    mMsgBoxEx - BAS Module
'Filename:  mMsgBoxEx.bas
'Author:    Jim Kahl
'Original:  Based on Ray Mercer's MsgBoxEx from www.shrinkwrapvb.com
'Purpose:   to Hook the standard Windows MsgBox so we can alter it to display the MsgBox
'           the way we want it to look with custom icons and custom button captions as
'           well as giving it a standard display position and time out period in the
'           event that the user does not respond within a certain period of time
'Modifications:
'   1 - Employs the use of a Resource File for possible Icons to be used
'   2 - Added HelpFile and Context parameter for direct support of MsgBox function
'       ie. changing any existing MsgBox statement to a MsgBoxEx statement will have
'       no effect on the operation of the application
'   3 - Changed Icon parameter to use Enumerated Constant Icon Indexes that are
'       stored in the resource file - this allows developer flexibility outside this
'       module to use predefined constants rather than try to remember the index numbers
'       or path names in trying to pass a valid Icon Handle to the routine
'   4 - Added ButtonText parameter to be passed in as a "Join"ed string using the "|"
'       character as a delimiter to change the caption of the buttons
'   5 - Added Timeout parameter to close window after desired Interval in Seconds - if
'       you want a 3 1/4 second timeout period, code the parameter as 3.25
'****************************************************************************************
'DO NOT CHANGE THIS MODULE UNLESS YOU KNOW WHAT YOU ARE DOING BECAUSE THE PROGRAM MAY
'CRASH UNEXPECTEDLY IF YOU DO SOMETHING INCORRECTLY
'****************************************************************************************
Option Explicit

'****************************************************************************************
'API CONSTANTS
'****************************************************************************************
Private Const WH_CBT As Long = 5
Private Const HCBT_ACTIVATE As Long = 5
Private Const HWND_TOP As Long = 0
'SetWindowPos
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_NOACTIVATE As Long = &H10
Private Const STM_SETICON As Long = &H170
Private Const SWVB_DEFAULT As Long = &HFFFFFFFF '-1 is reserved for centering
Private Const SWVB_CAPTION_DEFAULT As String = "SWVB_DEFAULT_TO_APP_TITLE"
'WindowMessages
Private Const WM_CLOSE = &H10
'OS Version
Private Const VER_PLATFORM_WIN32_WINDOWS = 1

'****************************************************************************************
'API TYPES
'****************************************************************************************
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

'****************************************************************************************
'API FUNCTIONS
'****************************************************************************************
Private Declare Function CallNextHookEx Lib "user32" ( _
                ByVal hHook As Long, _
                ByVal nCode As Long, _
                ByVal wParam As Long, _
                ByVal lParam As Long) _
                As Long
Private Declare Function FindWindowEx Lib "user32" _
        Alias "FindWindowExA" ( _
                ByVal ParenthWnd As Long, _
                ByVal ChildhWnd As Long, _
                ByVal ClassName As String, _
                ByVal Caption As String) _
                As Long
Private Declare Function GetClassName Lib "user32" _
        Alias "GetClassNameA" ( _
                ByVal hwnd As Long, _
                ByVal lpClassName As String, _
                ByVal nMaxCount As Long) _
                As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" _
        Alias "GetVersionExA" ( _
                lpVersionInformation As OSVERSIONINFO) _
                As Long
Private Declare Function GetWindowRect Lib "user32" ( _
                ByVal hwnd As Long, _
                lpRect As RECT) _
                As Long
Private Declare Function KillTimer Lib "user32" ( _
                ByVal hwnd As Long, _
                ByVal nIDEvent As Long) _
                As Long
Private Declare Function PostMessage Lib "user32" _
        Alias "PostMessageA" ( _
                ByVal hwnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                lParam As Any) _
                As Long
Private Declare Function SetTimer Lib "user32" ( _
                ByVal hwnd As Long, _
                ByVal nIDEvent As Long, _
                ByVal uElapse As Long, _
                ByVal lpTimerFunc As Long) _
                As Long
Private Declare Function SetWindowsHookEx Lib "user32" _
        Alias "SetWindowsHookExA" ( _
                ByVal idHook As Long, _
                ByVal lpfn As Long, _
                ByVal hmod As Long, _
                ByVal dwThreadId As Long) _
                As Long
Private Declare Function SetWindowPos Lib "user32" ( _
                ByVal hwnd As Long, _
                ByVal hWndInsertAfter As Long, _
                ByVal X As Long, _
                ByVal Y As Long, _
                ByVal cx As Long, _
                ByVal cy As Long, _
                ByVal wFlags As Long) _
                As Long
Private Declare Function SetWindowText Lib "user32.dll" _
        Alias "SetWindowTextA" ( _
                ByVal hwnd As Long, _
                ByVal lpString As String) _
                As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" ( _
                ByVal hHook As Long) _
                As Long

'****************************************************************************************
'ENUMERATED CONSTANTS
'****************************************************************************************
'MsgBoxEx
Public Enum pgaIconIndexes
    'each of these directly corresponds to the icon in the Resource file
    PicError = 101
    PicReport = 102
End Enum

'****************************************************************************************
'VARIABLES - PRIVATE
'****************************************************************************************
Private mhWnd As Long
Private mlTmrID As Long
Private msPrompt As String
Private mlHook As Long
Private mlLeft As Long
Private mlTop As Long
Private mhIcon As Long
Private mlOptions As Long
Private msButtonText() As String

'****************************************************************************************
'METHODS - PUBLIC
'****************************************************************************************
Public Function MsgBoxEx(ByVal Prompt As String, _
                Optional ByVal Options As VbMsgBoxStyle = vbOKOnly, _
                Optional ByVal Title As String = SWVB_CAPTION_DEFAULT, _
                Optional ByVal HelpFile As String, _
                Optional ByVal Context As Long, _
                Optional ByVal Left As Long = SWVB_DEFAULT, _
                Optional ByVal Top As Long = SWVB_DEFAULT, _
                Optional ByVal Icon As pgaIconIndexes, _
                Optional ByVal ButtonText As String = "||||", _
                Optional ByVal Timeout As Single = 0) As VbMsgBoxResult

    'Parameters:    Prompt - the message to appear in the MsgBox
    '               Options - the standard MsgBox options - if using a custom icon
    '                   do not set the icon in these options
    '               Title - the caption in the title bar of the MsgBox
    '               HelpFile - filespec of the help file to use when the Help button
    '                   is displayed
    '               Context - the help context ID of the Help topic to be displyed
    '               Left - the left position of the MsgBox in respect to the owner form
    '               Top - the left position of the MsgBox in respect to the owner form
    '               Icon - a valid handle to an Icon resource
    '               ButtonText - the text to be displayed on the buttons in order
    '                   in some cases the second and third elements of the string
    '                   are ignored, the fourth element always applys to the Help button
    '               Timeout - the time in seconds to dismiss the MsgBox - still has
    '                   millisecond precision ie. can be passed as 5.125
    'Returns:       standard MsgBox return values - ie. if using vbYesNoCancel and the
    '                   user clicks the second button the return value is vbNo - this
    '                   is regardless of what the caption states
    'NOTES:         The first 5 Parameters are the same as the standard MsgBox this is
    '                   so that the MsgBoxEx function can operate as a direct replacement
    '                   for current MsgBox statements - changing all MsgBox statements to
    '                   MsgBoxEx statements will not affect operation
    'Timeout Notes: Timeout does not work for vbAbortRetryIgnore or for vbYesNo since
    '               windows doesn't define a default return value - for vbOKOnly the Timeout
    '               will return vbOK for all others the Timeout returns vbCancel

    Dim hInst As Long
    Dim lThreadID As Long
    Dim picIco As StdPicture
    
    Set picIco = New StdPicture

    If ButtonText <> vbNullString Then
        'retrieve possible custom button captions
        msButtonText = Split(ButtonText, "|")
    End If

    hInst = App.hInstance
    lThreadID = GetCurrentThreadId()

    'First "subclass" the MsgBox function
    mlHook = SetWindowsHookEx(WH_CBT, AddressOf MsgBoxHook, hInst, lThreadID)

    'Save the new arguments as member variables to be used from the MsgBoxHook proc
    mlLeft = Left
    mlTop = Top
    mlOptions = Options
    msPrompt = Prompt

    'make sure that we are setting mhIcon to its proper value
    'either handle to an icon or 0
    If Icon <> 0& Then
        'this MUST be a valid icon - cannot be a bitmap
        Set picIco = LoadResPicture(Icon, vbResIcon)
        mhIcon = picIco.Handle
    Else
        mhIcon = 0&
    End If

    'default the MsgBox caption to app.title
    If Title = SWVB_CAPTION_DEFAULT Then
        Title = App.Title
    End If

    'if user wants custom icon make sure dialog has an icon to replace
    If mhIcon <> 0& Then
        'first we need to remove the existing icon
        mlOptions = mlOptions And &HFFFF8F
        'set the icon to a known value so we can replace it
        mlOptions = mlOptions Or vbCritical
    End If

    'set the timeout period
    If Timeout >= 0.001 Then
        Timeout = Timeout * 1000
        mlTmrID = SetTimer(0&, 0&, Timeout, AddressOf TimerProc)
    End If

    'show the MsgBox and let the hook process take care of the rest
    MsgBoxEx = MsgBox(msPrompt, mlOptions, Title, HelpFile, Context)

End Function

'****************************************************************************************
'METHODS - PRIVATE
'****************************************************************************************
Private Function MsgBoxHook( _
        ByVal nCode As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

    'performs the hook of the MsgBox function to allow our custom changes to take effect
    Dim lHgt As Long
    Dim lWid As Long
    Dim lSize As Long
    Dim rcMsgBox As RECT
    Dim sBuff As String
    Dim lPosX As Long
    Dim lPosY As Long
    Dim X As Long
    Dim Y As Long
    Dim hwnd As Long

    'NOTE: Class names we are interested in are as follows
    '#32770 - cbt dialog
    'Static - Icon
    'Button - Command Button
    '    Debug.Print "hook proc called"

    'Call next hook in the chain and return the value
    '(this is the polite way to allow other hhoks to function too)
    MsgBoxHook = CallNextHookEx(mlHook, nCode, wParam, lParam)

    'hook only the activate msg
    If nCode = HCBT_ACTIVATE Then
        'handle only standard MsgBox class windows
        'this is the most efficient method to allocate strings in VB
        'according to Brad Martinez's results with tools from NuMega
        sBuff = Space$(32)

        'GetClassName will truncate the class name if it doesn't fit in the buffer
        'we only care about the first 6 chars anyway
        lSize = GetClassName(wParam, sBuff, 32)
        If Left$(sBuff, lSize) <> "#32770" Then
            'not a standard msgBox
            'we can just quit because we already called CallNextHookEx
            Exit Function
        End If

        'get the handle for the MsgBox
        mhWnd = FindWindowEx(0&, 0&, "#32770", vbNullString)
        '        Debug.Print "MsgBox handle " & mhWnd

        'store MsgBox window size in case we need it
        Call GetWindowRect(wParam, rcMsgBox)

        'get size of msgbox
        lHgt = (rcMsgBox.Bottom - rcMsgBox.Top) / 2
        lWid = (rcMsgBox.Right - rcMsgBox.Left) / 2

        'store parent window size
        Call GetWindowRect(GetParent(wParam), rcMsgBox)

        'center msgbox to parent form
        lPosY = rcMsgBox.Top + (rcMsgBox.Bottom - rcMsgBox.Top) / 2
        lPosX = rcMsgBox.Left + (rcMsgBox.Right - rcMsgBox.Left) / 2

        'if user passed in specific values then use those instead
        If mlLeft = SWVB_DEFAULT Then 'default
            X = lPosX - lWid
        Else
            X = mlLeft
        End If

        If mlTop = SWVB_DEFAULT Then 'default
            Y = lPosY - lHgt
        Else
            Y = mlTop
        End If

        'If user passed in custom icon use that instead of the standard Windows icon
        If mhIcon <> 0& Then
            hwnd = FindWindowEx(wParam, 0&, "Static", vbNullString)
            '            Debug.Print "Icon Handle " & hWnd
            Call SendMessage(hwnd, STM_SETICON, mhIcon, ByVal 0&)
        End If

        'handle to the message window
        hwnd = FindWindowEx(wParam, 0&, "Static", msPrompt)

        'Manually set the MsgBox window position before Windows shows it
        SetWindowPos wParam, HWND_TOP, X, Y, 0&, 0&, SWP_NOZORDER + SWP_NOACTIVATE + SWP_NOSIZE

        'change the captions of the buttons
        SetButtonCaptions wParam

ErrHandler:
        'unhook the dialog and we are out clean!
        UnhookWindowsHookEx mlHook
        '        Debug.Print "unhook"
    End If

End Function

Private Sub SetButtonCaptions(ByVal wParam As Long)
    'sets the captions of the buttons to developer defined values
    Dim hwnd As Long
    Dim sText As String

    'check the various buttons and change their caption to our passed string parts

    'Special note: when attempting to get the handle of buttons with captions
    'Yes, No, Abort, Retry or Ignore with FindWindowEx what we are actually
    'looking for is &Yes, &No, etc. the captions are shown that way in Win95 but
    'not in XP, not sure about other OS Versions

    'Retry/Cancel - note that the third element of the string is ignored
    If (mlOptions And vbRetryCancel) = vbRetryCancel Then
            ButtonCaption wParam, msButtonText(0), "&Retry"
            ButtonCaption wParam, msButtonText(0), "Cancel"
        'Yes/No - note that the third element of the string is ignored
    ElseIf (mlOptions And vbYesNo) = vbYesNo Then
            ButtonCaption wParam, msButtonText(0), "&Yes"
            ButtonCaption wParam, msButtonText(0), "&No"
        'Yes/No/Cancel
    ElseIf (mlOptions And vbYesNoCancel) = vbYesNoCancel Then
            ButtonCaption wParam, msButtonText(0), "&Yes"
            ButtonCaption wParam, msButtonText(0), "&No"
            ButtonCaption wParam, msButtonText(0), "Cancel"
        'Abort/Retry/Ignore
    ElseIf (mlOptions And vbAbortRetryIgnore) = vbAbortRetryIgnore Then
            ButtonCaption wParam, msButtonText(0), "&Abort"
            ButtonCaption wParam, msButtonText(0), "&Retry"
            ButtonCaption wParam, msButtonText(0), "&Ignore"
        'OK/Cancel - note that the third element of the string is ignored in this case
    ElseIf (mlOptions And vbOKCancel) = vbOKCancel Then
            ButtonCaption wParam, msButtonText(0), "&Ok"
            ButtonCaption wParam, msButtonText(0), "Cancel"
        'OK Only - note that the second and third element of the string is ignored
    ElseIf (mlOptions And vbOKOnly) = vbOKOnly Then
            ButtonCaption wParam, msButtonText(0), "Ok"
    End If

    'because the Help button only deals with help files it doesn't have a return
    'value but rather brings up the help file and context that are defined in the
    'call - we can still change the caption of the button but we need to test for
    'it by itself

    If (mlOptions And vbMsgBoxHelpButton) = vbMsgBoxHelpButton Then
        If msButtonText(3) <> vbNullString Then
            'Note Win95 need to look for &Help but XP need to look for Help so we
            'need to determine what OS version we are using - not sure about other
            'OS versions
            If IsWin95 Then
                sText = "&Help"
            Else
                sText = "Help"
            End If
            hwnd = FindWindowEx(wParam, 0&, "Button", sText)
            SetWindowText hwnd, msButtonText(3)
        End If
    End If

End Sub

Private Sub ButtonCaption(ByVal wParam As Long, ByVal TxtBut As String, ByVal StrCap As String)
    If LenB(TxtBut) Then
        SetWindowText FindWindowEx(wParam, 0&, "Button", StrCap), TxtBut
    End If
End Sub

Private Function TimerProc( _
        ByVal hwnd As Long, _
        ByVal uMsg As Long, _
        ByVal idEvent As Long, _
        ByVal dwTime As Long) As Long

    Dim lRet As Long

    'this is useless for vbYesNo and vbAbortRetryIgnore message boxes
    lRet = PostMessage(mhWnd, WM_CLOSE, 0, ByVal 0&)

    'Kill timer.
    If mlTmrID Then
        Call KillTimer(0&, mlTmrID)
        mlTmrID = 0
    End If
End Function

Private Function IsWin95() As Boolean
    Static osv As OSVERSIONINFO
    Static bRet As Boolean

    'just do this once
    If osv.dwPlatformId = 0 Then
        osv.dwOSVersionInfoSize = Len(osv)
        Call GetVersionEx(osv)
        'we can check the Major and Minor versions if necessary for determining what
        'the help button caption is
        bRet = (osv.dwMinorVersion < 10) And _
            (osv.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS)
    End If
    IsWin95 = bRet
End Function


