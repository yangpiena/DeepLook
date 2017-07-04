Attribute VB_Name = "ModRegistry"
'*************************************************************************
''               人人为我，我为人人
''      枕善居VB及.NET源码博客汉化收藏整理
''网    站：http://www.Mndsoft.com/
''e-mail  ：mndsoft@126.com
''发布日期：2009-10-08 10:11:44
''QQ      ：88382850
''   如果您有新的、好的代码可以提供给枕善居上发布，让大家学习哦!
''----------------------------------------------------------------------
'**系统名称：VB及.NET工程源代码扫描分析工具 V4.12.0
'**模块描述：
'**模 块 名：ModRegistry
'**创 建 人：
'**汉 化 者：枕善居(mndsoft)
'**日    期：2009-10-08 10:11:44
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V4.12.0
'*************************************************************************
' THIS MODULE WAS NOT WRITTEN BY DEAN CAMERA. IT'S ORIGINAL AUTHOR IS UNKNOWN.
'             I CANNOT OFFER ANY SUPPORT FOR THIS MODULE.

' Registry Subroutines - for getting Component names, etc.

Option Explicit

Global Const REG_SZ As Long = 1
Global Const REG_DWORD As Long = 4

Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003

Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259

Global Const KEY_ALL_ACCESS = &H3F

Global Const REG_OPTION_NON_VOLATILE = 0

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey& Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String)
Private Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String)

Public Function DeleteKey(lPredefinedKey As Long, sKeyName As String)
    ' Description:
    '   This Function will Delete a key
    '
    ' Syntax:
    '   DeleteKey Location, KeyName
    '
    '   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
    '   , HKEY_USERS
    '
    '   KeyName is name of the key you wish to delete, it may include subkeys (example "Key1\SubKey1")


    Dim lRetVal As Long         'result of the SetValueEx function

    'open the specified key
    lRetVal = RegDeleteKey(lPredefinedKey, sKeyName)
End Function

Public Function DeleteValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
    ' Description:
    '   This Function will delete a value
    '
    ' Syntax:
    '   DeleteValue Location, KeyName, ValueName
    '
    '   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
    '   , HKEY_USERS
    '
    '   KeyName is the name of the key that the value you wish to delete is in
    '   , it may include subkeys (example "Key1\SubKey1")
    '
    '   ValueName is the name of value you wish to delete

    Dim lRetVal As Long         'result of the SetValueEx function
    Dim hKey As Long         'handle of open key

    'open the specified key

    lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = RegDeleteValue(hKey, sValueName)
    RegCloseKey (hKey)
End Function

Public Function SetValueEx(ByVal hKey As Long, sValueName As String, LType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String

    Select Case LType
        Case REG_SZ
            sValue = vValue
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, LType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, LType, lValue, 4)
    End Select
End Function

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim LType As Long
    Dim lValue As Long
    Dim sValue As String

    On Local Error GoTo QueryValueExError

    ' Determine the size and type of data to be read

    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, LType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5

    Select Case LType
            ' For strings
        Case REG_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, LType, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch)
            Else
                vValue = Empty
            End If

            ' For DWORDS
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, LType, lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
            'all other data types not supported
            lrc = -1
    End Select

QueryValueExExit:

    QueryValueEx = lrc
    Exit Function

QueryValueExError:

    Resume QueryValueExExit
End Function

Public Function CreateNewKey(lPredefinedKey As Long, sNewKeyName As String)
    ' Description:
    '   This Function will create a new key
    '
    ' Syntax:
    '   QueryValue Location, KeyName
    '
    '   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
    '   , HKEY_USERS
    '
    '   KeyName is name of the key you wish to create, it may include subkeys (example "Key1\SubKey1")



    Dim hNewKey As Long 'handle to the new key

     RegCreateKeyEx lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, 0

    RegCloseKey (hNewKey)
End Function

Public Function SetKeyValue(lPredefinedKey As Long, sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
    ' Description:
    '   This Function will set the data field of a value
    '
    ' Syntax:
    '   QueryValue Location, KeyName, ValueName, ValueSetting, ValueType
    '
    '   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
    '   , HKEY_USERS
    '
    '   KeyName is the key that the value is under (example: "Key1\SubKey1")
    '
    '   ValueName is the name of the value you want create, or set the value of (example: "ValueTest")
    '
    '   ValueSetting is what you want the value to equal
    '
    '   ValueType must equal either REG_SZ (a string) Or REG_DWORD (an integer)

    Dim lRetVal As Long 'result of the SetValueEx function
    Dim hKey As Long    'handle of open key

    'open the specified key

    lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
    RegCloseKey (hKey)
End Function

Public Function QueryValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
    ' Description:
    '   This Function will return the data field of a value
    '
    ' Syntax:
    '   Variable = QueryValue(Location, KeyName, ValueName)
    '
    '   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
    '   , HKEY_USERS
    '
    '   KeyName is the key that the value is under (example: "Software\Microsoft\Windows\CurrentVersion\Explorer")
    '
    '   ValueName is the name of the value you want to access (example: "link")

    Dim lRetVal As Long         'result of the API functions
    Dim hKey As Long         'handle of opened key
    Dim vValue As Variant      'setting of queried value


    lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = QueryValueEx(hKey, sValueName, vValue)
    'MsgBox vValue
    QueryValue = vValue
    RegCloseKey (hKey)
End Function
