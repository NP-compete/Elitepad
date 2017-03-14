Attribute VB_Name = "modMRU"
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

Public Sub GetMRUList()
    '** Description:
    '** Get MRU List from the registry
    On Error GoTo GetMRUError
    Dim I As Integer
    Dim FileName As String
    
    Set frmMDI.MRUList = New Collection 'Create new collection
    For I = 1 To 8
        FileName = RGGetKeyValue(HKEY_LOCAL_MACHINE, MRUPath, "Name" & I) 'Get file name
        If Len(FileName) > 2 Then
            frmMDI.MRUList.Add FileName 'Add file name to collection
        End If
    Next
    ShowMRUList 'Call DisplayMRUList sub
GetMRUError:
    ErrorLog "modMRU\GetMRUList"
End Sub

Public Sub ShowMRUList()
    '** Description:
    '** Show MRU List
    On Error GoTo ShowMRUError
    Dim I As Integer
    
    For I = 1 To 8
        If I > frmMDI.MRUList.Count Then Exit For
        ' Set menu caption
        frmMDI.mnuFileMRUItem(I).Caption = I & ". " & GetFTitle(frmMDI.MRUList(I))
        ' Set menu tag to file name
        frmMDI.mnuFileMRUItem(I).Tag = frmMDI.MRUList(I)
        ' Show menu
        frmMDI.mnuFileMRUItem(I).Visible = True
    Next
    
    For I = frmMDI.MRUList.Count + 1 To 8
        frmMDI.mnuFileMRUItem(I).Visible = False 'Hide empty menus
    Next
    
    If frmMDI.MRUList.Count > 0 Then
        frmMDI.mnuFileMRUTemp.Visible = False 'Hide temporary menu
    End If
ShowMRUError:
    ErrorLog "modMRU\ShowMRUList"
End Sub

Public Sub SaveMRUFile(ByVal FileName As String)
    '** Description:
    '** Save file in the registry
    On Error GoTo MRUSaveError
    Dim I As Integer

    For I = 1 To 8
        If I > frmMDI.MRUList.Count Then Exit For
        If LCase(frmMDI.MRUList(I)) = LCase(FileName) Then 'If filename exist in the
            frmMDI.MRUList.Remove I                     'collection exit sub
            Exit For
        End If
    Next I
    
    If frmMDI.MRUList.Count > 0 Then 'If the collection is not empty
        frmMDI.MRUList.Add FileName, , 1 'add file to begining of the collecton
    Else 'else
        frmMDI.MRUList.Add FileName 'just add it
    End If
    
    If frmMDI.MRUList.Count > 8 Then 'If there are more items than 8 remove the last one
        frmMDI.MRUList.Remove 9
    End If
    
    For I = 1 To 8
        If I > frmMDI.MRUList.Count Then 'If no more files then leave it empty
            FileName = ""
        Else 'else
            FileName = frmMDI.MRUList(I) 'add it
        End If
        ' Add file to the registry
        RGSetKeyValue HKEY_LOCAL_MACHINE, MRUPath, "Name" & I, FileName
    Next I
    GetMRUList
MRUSaveError:
    ErrorLog "modMRU\SaveMRUFile"
End Sub
