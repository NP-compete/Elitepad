Attribute VB_Name = "modFileOperations"
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

Private Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer

Private Type SHFILEOPSTRUCT
        hWnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAnyOperationsAborted As Long
        hNameMappings As Long
        lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3
Private Const FO_RENAME = &H4
Private Const FOF_ALLOWUNDO = &H40

Public Function Associate(Program As String, Extension As String, Description As String, Optional Icon As String)
    '** Description:
    '** Associate file with ElitePad
    RGCreateKey HKEY_CLASSES_ROOT, "." & Extension
    RGSetKeyValue HKEY_CLASSES_ROOT, "." & Extension, "", Extension & "file"
    
    RGCreateKey HKEY_CLASSES_ROOT, Extension & "file"
    RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\shell"
    If LCase(Extension) = "bat" Then
        RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\shell\edit"
        RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\shell\edit\command"
        RGSetKeyValue HKEY_CLASSES_ROOT, Extension & "file\shell\edit\command", "", Program & " " & "%1" 'Set file path
    Else
        RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\shell\open"
        RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\shell\open\command"
        RGSetKeyValue HKEY_CLASSES_ROOT, Extension & "file\shell\open\command", "", Program & " " & "%1" 'Set file path
    End If
    RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\DefaultIcon"
    
    RGSetKeyValue HKEY_CLASSES_ROOT, Extension & "file", "", Description 'Set file description
    RGSetKeyValue HKEY_CLASSES_ROOT, Extension & "file\DefaultIcon", "", Icon 'Set file icon
End Function

Public Function GetFTitle(strFilename As String)
    '** Description:
    '** Get file title from file name
    On Error GoTo GFTError
    Dim cbBuf As String
    
    cbBuf = String(250, vbNullChar) 'Fill buffer with null chars
    GetFileTitle strFilename, cbBuf, Len(cbBuf) 'Get file title
    GetFTitle = Left(cbBuf, InStr(1, cbBuf, vbNullChar) - 1) 'Extract file title from buffer
GFTError:
    ErrorLog "modFileOperations/GetFTitle"
End Function

Public Function FormatSize(ByVal Amount As Long) As String
    '** Description:
    '** Format file size
    Dim Buffer As String
    Dim Result As String
    
    Buffer = Space$(255) 'Fill buffer
    Result = StrFormatByteSize(Amount, Buffer, Len(Buffer)) 'Format file size
    If InStr(Result, vbNullChar) > 1 Then
        FormatSize = Left$(Result, InStr(Result, vbNullChar) - 1)
    End If
End Function

Public Function CopyFile(sFrom As String, sTo As String)
    '** Description:
    '** Copy file
    On Error GoTo CopyError
    Dim SHFO As SHFILEOPSTRUCT
    With SHFO
        .wFunc = FO_COPY 'Set copy metod
        .pFrom = sFrom 'Set from filename
        .pTo = sTo 'Set to filename
    End With
    SHFileOperation SHFO 'Copy file
CopyError:
    ErrorLog "modFileOperations\CopyFile"
End Function

Public Function DeleteFile(sFileName As String)
    '** Description:
    '** Delete file
    On Error GoTo DeleteError
    Dim SHFO As SHFILEOPSTRUCT
    With SHFO
        .wFunc = FO_DELETE 'Set delete metod
        .pFrom = sFileName 'Set from filename
        .fFlags = FOF_ALLOWUNDO 'Allow undo
    End With
    SHFileOperation SHFO 'Delete file
DeleteError:
    ErrorLog "modFileOperations\DeleteFile"
End Function

Public Function RenameFile(sFrom As String, sTo As String)
    '** Description:
    '** Rename file
    On Error GoTo RenameError
    Dim SHFO As SHFILEOPSTRUCT
    With SHFO
        .wFunc = FO_RENAME 'Set delete method
        .pFrom = sFrom 'Set from filename
        .pTo = sTo 'Set to filename
    End With
    SHFileOperation SHFO 'Rename file
RenameError:
    ErrorLog "modFileOperations\RenameFile"
End Function

Public Function FileExists(sFileName As String) As Boolean
    '** Description:
    '** Check to see if file exists
    On Error GoTo FExistsError
    Dim F As String
    F = FreeFile
    Open sFileName For Input As #F 'Open file
    Close #F
FExistsError:
    If Err.Number = 53 Then 'If doesn't exists
        FileExists = False 'Set FileExists to False
    ElseIf Err.Number = 0 Then 'else if exists
        FileExists = True 'Set FileExists to True
    End If
End Function
