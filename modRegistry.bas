Attribute VB_Name = "modRegistry"
    '***************************************************************'
    '                         ELITEPAD                              '
    '                        Written by                             '
    '                       Andrea Batina                           '
    '                                                               '
    '  You are free to use the source code in your private,         '
    '  non-commercial, projects without permission. If you want     '
    '  to use this code in commercial projects EXPLICIT permission  '
    '  from the author is required.                                 '
    '                                                               '
    '                                                               '
    '               Copyright © Andrea Batina 1999-2000             '
    '***************************************************************'

Option Explicit

' Win32 declarations for the registry access
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE And KEY_ENUMERATE_SUB_KEYS And KEY_NOTIFY And KEY_CREATE_SUB_KEY And KEY_CREATE_LINK And KEY_SET_VALUE
Private Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
Private Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
Private Const REG_SZ = 1                        ' Unicode nul terminated string

Private Const ERROR_SUCCESS = 0&

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003

Public cEnumValues As Collection
Public cEnumKeys As Collection

Public Function RGCreateKey(hKey As Long, SubKey As String)
    '** Description:
    '** Create a new key
    '**
    '** Syntax:
    '** RGCreateKey(hKey,SubKey)
    '**
    '** Example:
    '** RGCreateKey(HKEY_LOCAL_MACHINE,"Software\VBReality")
    Dim lngRet As Long
    Dim lngResult As Long
    Dim lngDis As Long
    
    ' Create a new key
    lngRet = RegCreateKeyEx(hKey, SubKey, 0&, 0&, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, lngResult, lngDis)
    lngRet = RegCloseKey(lngResult) 'Close key
End Function

Public Function RGDeleteKey(hKey As Long, SubKey As String)
    '** Description:
    '** Delete key
    '**
    '** Syntax:
    '** RGDeleteKey(hKey,SubKey)
    '**
    '** Example:
    '** RGDeleteKey(HKEY_LOCAL_MACHINE,"Software\VBReality")
    RegDeleteKey hKey, SubKey 'Delete key
End Function

Public Function RGSetKeyValue(hKey As Long, SubKey As String, ValueName As String, sValue As String)
    '** Description:
    '** Set key value
    '**
    '** Syntax:
    '** RGSetKeyValue(hKey,SubKey,ValueName,ValueSetting)
    '**
    '** Example:
    '** RGSetKeyValue(HKEY_LOCAL_MACHINE,"Software\VBReality","Written by","Andrea Batina")
    Dim lngRet As Long
    Dim lngResult As Long
    
    ' Open key
    lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
    If lngRet = ERROR_SUCCESS Then 'If key exist
        ' Set key value
        RegSetValueEx lngResult, ValueName, 0, REG_SZ, ByVal sValue, Len(sValue)
        RegFlushKey lngResult 'Update registry
        RegCloseKey lngResult 'Close key
    End If
End Function

Public Function RGGetKeyValue(hKey As Long, SubKey As String, ValueName As String, Optional Default As String = "")
    '** Description:
    '** Get key value
    '**
    '** Syntax:
    '** RGGetKeyValue(hKey,SubKey,ValueName,Default)
    '**
    '** Example:
    '** RGGetKeyValue(HKEY_LOCAL_MACHINE,"Software\VBReality","Written by","ME")
    Dim lngRet As Long
    Dim lngResult As Long
    Dim sData As String
    
    ' Open key
    lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
    If lngRet = ERROR_SUCCESS Then 'If key exist
        sData = String(128, vbNullChar) 'Fill buffer with null chars
        ' Get key value
        lngRet = RegQueryValueEx(lngResult, ValueName, 0, REG_SZ, ByVal sData, Len(sData))
        ' If valuename doesnt exist set default value
        If Not lngRet = ERROR_SUCCESS Then RGGetKeyValue = Default: Exit Function
        RGGetKeyValue = Left(sData, InStr(1, sData, vbNullChar) - 1)
        RegCloseKey lngResult 'Close key
    Else 'If key doesnt exist
        RGGetKeyValue = Default 'Set default value
    End If
End Function

Public Function RGDeleteKeyValue(hKey As Long, SubKey As String, ValueName As String)
    '** Description:
    '** Delete key value
    '**
    '** Syntax:
    '** RGDeleteKeyValue(hKey,SubKey,ValueName)
    '**
    '** Example:
    '** RGDeleteKeyValue(HKEY_LOCAL_MACHINE,"Software\VBReality","Written by")
    Dim lngRet As Long
    Dim lngResult As Long
       
    ' Open key
    lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
    RegDeleteValue lngResult, ValueName 'Delete key value
End Function

Public Function RGEnumKeyValues(hKey As Long, SubKey As String)
    '** Description:
    '** Enum key values
    '**
    '** Syntax:
    '** RGEnumKeyValues(hKey,SubKey)
    '**
    '** Example:
    '** RGEnumKeyValues(HKEY_LOCAL_MACHINE,"Software\VBReality")
    '** For i = 1 to cEnumValues.count
    '**     List1.AddItem cEnumValues.Item(i)
    '** Next
    Dim lngRet As Long
    Dim lngResult As Long
    Dim sData As String
    Dim intIndex As Integer

    ' Open key
    lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
    If lngRet = ERROR_SUCCESS Then 'If key exist
        Set cEnumValues = New Collection 'Make new collection
        Do
            sData = String(128, vbNullChar) 'Fill buffer with null chars
            ' Enum key values
            lngRet = RegEnumValue(lngResult, intIndex, sData, Len(sData), 0, ByVal 0&, ByVal 0&, ByVal 0&)
            If lngRet <> 0 Then Exit Do 'If there are no more values exit do
            cEnumValues.Add Left(sData, InStr(1, sData, vbNullChar) - 1) 'Add values
            intIndex = intIndex + 1 'Increase counter by 1
        Loop
        RegCloseKey lngResult 'Close key
    End If
End Function

Public Function RGEnumKeys(hKey As Long, SubKey As String)
    '** Description:
    '** Enum keys
    '**
    '** Syntax:
    '** RGEnumKeys(hKey,SubKey)
    '**
    '** Example:
    '** RGEnumKeys(HKEY_LOCAL_MACHINE,"Software\VBReality")
    '** For i = 1 to cEnumKeys.count
    '**     List1.AddItem cEnumKeys.Item(i)
    '** Next
    Dim lngRet As Long
    Dim lngResult As Long
    Dim sData As String
    Dim intIndex As Integer

    ' Open key
    lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
    If lngRet = ERROR_SUCCESS Then 'If key exist
        Set cEnumKeys = New Collection 'Make new collection
        Do
            sData = String(128, vbNullChar) 'Fill buffer with null chars
            ' Enum key keys
            lngRet = RegEnumKey(lngResult, intIndex, sData, Len(sData))
            If lngRet <> 0 Then Exit Do 'If there are no more keys exit do
            cEnumKeys.Add Left(sData, InStr(1, sData, vbNullChar) - 1) 'Add keys
            intIndex = intIndex + 1 'Increase counter by 1
        Loop
        RegCloseKey lngResult 'Close key
    End If
End Function
