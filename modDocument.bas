Attribute VB_Name = "modDocument"
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

' Filter
Public Const epFilter = "Text Files (*.txt)|*.txt|RichText Files (*.rtf)|*.rtf|Log Files (*.log)|*.log|Batch Files (*.bat)|*.bat|INI Files (*.ini)|*.ini|All Files (*.*)|*.*|"

Public Function CreateNewDocument()
    '** Description:
    '** Create a new document
    On Error GoTo NewError
    Static DocCount As Long
    Dim frmDoc As frmDocument
    
    Set frmDoc = New frmDocument 'Create new form
    DocCount = DocCount + 1 'Increase document counter
    frmDoc.Caption = "Document " & DocCount 'Set document caption
    frmDoc.Show 'Show document
NewError:
    ErrorLog "modDocument/CreateNewDocument"
End Function

Public Function OpenDocument()
    '** Description:
    '** Open an existing document
    On Error GoTo OpenError
    Dim I As Integer
    Dim fType As String
    
    With frmMDI
        .CmDlg.CancelError = True 'Set cancel error to true
        .CmDlg.MultiSelect = True 'Allow multi select
        .CmDlg.DialogTitle = "Select file(s) to open" 'Set dialog title
        .CmDlg.Filter = epFilter 'Set file filter
        .CmDlg.FilterIndex = 1 'Set filter index
        .CmDlg.ShowOpen 'Show open dialog
        
        For I = 1 To .CmDlg.cFileName.Count 'Open all selected files
            ' Get file extension
            If UCase(Right(.CmDlg.cFileTitle(I), 3)) = "RTF" Then
                fType = rtfText
            Else
                fType = rtfText
            End If
            CreateNewDocument
            .ActiveForm.rtfText.LoadFile .CmDlg.cFileName(I), fType 'Load file
            .ActiveForm.Caption = .CmDlg.cFileName(I) 'Set form caption
            .ActiveForm.bChanged = False 'Set bChanged flag to false
            SaveMRUFile .CmDlg.cFileName(I) 'Save file into MRU List
        Next
    End With
OpenError:
    If Err.Number = 32755 Then Exit Function 'If canceled then exit function
    ErrorLog "modDocument/OpenDocument"
End Function

Public Function SaveDocument()
    '** Description:
    '** Save the active document
    On Error GoTo SaveError
    Dim fType As String
    
    With frmMDI.ActiveForm
        If Left(.Caption, 8) = "Document" Then
            'If it is not saved then call SaveDocumentAs function
            SaveDocumentAs
        Else
            ' Get file extension
            If UCase(Right(.Caption, 3)) = "RTF" Then
                fType = rtfText
            Else
                fType = rtfText
            End If
            .rtfText.SaveFile .Caption, fType 'Save document
            .bChanged = False 'Set bChanged flag to false
        End If
    End With
SaveError:
    ErrorLog "modDocument/SaveDocument"
End Function

Public Function SaveDocumentAs()
    '** Description:
    '** Save the active document with a new name
    On Error GoTo SaveAsError
    Dim fType As String
    
    With frmMDI
        .CmDlg.CancelError = True 'Set cancel error to true
        .CmDlg.DialogTitle = "Save As" 'Set dialog title
        .CmDlg.Filter = epFilter 'Set file filter
        .CmDlg.FilterIndex = 1 'Set filter index
        .CmDlg.DefaultFilename = GetFTitle(.ActiveForm.Caption) 'Get file title and put it in common dialog as default filename
        .CmDlg.ShowSave 'Show save dialog
    
        ' Get file extension
        If UCase(Right(.CmDlg.cFileTitle(1), 3)) = "RTF" Then
            fType = rtfText
        Else
            fType = rtfText
        End If
        
        .ActiveForm.rtfText.SaveFile .CmDlg.cFileName(1), fType 'Save document
        .ActiveForm.Caption = .CmDlg.cFileName(1) 'Set form caption
        .ActiveForm.bChanged = False 'Set bChanged flag to false
        SaveMRUFile .CmDlg.cFileName(1) 'Save file into MRU List
    End With
SaveAsError:
    If Err.Number = 32755 Then Exit Function 'If canceled then exit function
    ErrorLog "modDocument/SaveDocumentAs"
End Function

Public Function SaveAllDocuments()
    '** Description:
    '** Save all changed documents
    On Error GoTo SaveAllError
    Dim fType As String
    
    While Forms.Count > 1
        With frmMDI
            If .ActiveForm.bChanged = True Then
                If Left(.ActiveForm.Caption, 8) = "Document" Then
                    .CmDlg.DialogTitle = "Save As" 'Set dialog title
                    .CmDlg.Filter = epFilter 'Set file filter
                    .CmDlg.CancelError = False 'Set cancel error to false
                    .CmDlg.DefaultFilename = .ActiveForm.Caption 'Set default filename
                    .CmDlg.ShowSave 'Show save dialog
                    
                    ' Get file extension
                    If UCase(Right(.CmDlg.cFileTitle(1), 3)) = "RTF" Then
                        fType = rtfText
                    Else
                        fType = rtfText
                    End If
            
                    .ActiveForm.rtfText.SaveFile .CmDlg.cFileName(1), fType 'Save document
                    .ActiveForm.Caption = .CmDlg.cFileName(1) 'Set form caption
                    .ActiveForm.bChanged = False 'Set bChanged flag to false
                    Unload .ActiveForm
                Else
                    ' Get file extension
                    If UCase(Right(.CmDlg.cFileTitle(1), 3)) = "RTF" Then
                        fType = rtfText
                    Else
                        fType = rtfText
                    End If
                    .ActiveForm.rtfText.SaveFile .ActiveForm.Caption, fType 'Save document
                    .ActiveForm.bChanged = False 'Set bChanged flag to false
                    Unload .ActiveForm
                End If
            End If
        End With
    Wend
SaveAllError:
    ErrorLog "modDocument/SaveAllDocuments"
End Function

Public Function SaveSelectionAs()
    '** Description:
    '** Save selection as file
    On Error GoTo SaveSelError
    Dim fType As String
    
    With frmMDI
        .rtfTemp.Text = "" 'Empty hidden richtextbox
        SendMessage .ActiveForm.rtfText.hWnd, WM_COPY, 0, 0& 'Copy selected text from rtfText
        SendMessage .rtfTemp.hWnd, WM_PASTE, 0, 0& 'and paste it into rtfTemp
        
        .CmDlg.DialogTitle = "Save Selection As" 'Set dialog title
        .CmDlg.Filter = epFilter 'Set file filter
        .CmDlg.FilterIndex = 1 'Set filter index
        .CmDlg.DefaultFilename = GetFTitle(.ActiveForm.Caption) 'Set default filename
        .CmDlg.ShowSave 'Show save dialog
        
        ' Get file extension
        If UCase(Right(.CmDlg.cFileTitle(1), 3)) = "RTF" Then
            fType = rtfText
        Else
            fType = rtfText
        End If
    
        .rtfTemp.SaveFile .CmDlg.cFileName(1), fType 'Save document
    End With
SaveSelError:
    ErrorLog "modDocument\SaveSelectionAs"
End Function

Public Function RevertDocument()
    '** Description:
    '** Revert the current document to the last saved version
    On Error GoTo RevertError
    Dim fType As String
    
    With frmMDI.ActiveForm
        ' If it isn't saved then exit function
        If Left(.Caption, 8) = "Document" Then Exit Function
    
        If .bChanged = True Then 'If document is changed
            ' Ask to revert document to last saved version
            MsgE "Do you want to revert [" & .Caption & "] to last saved version ?", "ElitePad - Revert", 1, False
            If frmMDI.bYes = True Then  'Save document
                ' Get file extension
                If UCase(Right(.Caption, 3)) = "RTF" Then
                    fType = rtfText
                Else
                    fType = rtfText
                End If
                .rtfText.LoadFile .Caption, fType 'Load document
                .bChanged = False 'Set bChanged flag to false
            Else 'No
                Exit Function
            End If
        End If
    End With
RevertError:
    ErrorLog "modDocument\RevertDocument"
End Function

Public Function Bold()
    If frmMDI.ActiveForm Is Nothing Then Exit Function
    With frmMDI.ActiveForm
        If .rtfText.SelBold = True Then 'If it is bold
            .rtfText.SelBold = False 'Make it normal
            frmMDI.tbrFormat.Buttons("Bold").Value = tbrUnpressed 'Set button to unpressed
        Else 'If it is not bold
            .rtfText.SelBold = True 'Make it
            frmMDI.tbrFormat.Buttons("Bold").Value = tbrPressed 'Set button to pressed
        End If
    End With
End Function

Public Function Italic()
    If frmMDI.ActiveForm Is Nothing Then Exit Function
    With frmMDI.ActiveForm
        If .rtfText.SelItalic = True Then 'If it is italic
            .rtfText.SelItalic = False 'Make it normal
            frmMDI.tbrFormat.Buttons("Italic").Value = tbrUnpressed 'Set button to unpressed
        Else 'If it is not italic
            .rtfText.SelItalic = True 'Make it
            frmMDI.tbrFormat.Buttons("Italic").Value = tbrPressed 'Set button to pressed
        End If
    End With
End Function

Public Function Underline()
    If frmMDI.ActiveForm Is Nothing Then Exit Function
    With frmMDI.ActiveForm
        If .rtfText.SelUnderline = True Then 'If it is underline
            .rtfText.SelUnderline = False 'Make it normal
            frmMDI.tbrFormat.Buttons("Underline").Value = tbrUnpressed 'Set button to unpressed
        Else 'If it is not underline
            .rtfText.SelUnderline = True 'Make it
            frmMDI.tbrFormat.Buttons("Underline").Value = tbrPressed 'Set button to pressed
        End If
    End With
End Function

Public Function Strikethru()
    If frmMDI.ActiveForm Is Nothing Then Exit Function
    With frmMDI.ActiveForm
        If .rtfText.SelStrikeThru = True Then 'If it is strikethru
            .rtfText.SelStrikeThru = False 'Make it normal
            frmMDI.tbrFormat.Buttons("StrikeThru").Value = tbrUnpressed 'Set button to unpressed
        Else 'If it is not strikethru
            .rtfText.SelStrikeThru = True 'Make it
            frmMDI.tbrFormat.Buttons("StrikeThru").Value = tbrPressed 'Set button to pressed
        End If
    End With
End Function

Public Function AlignLeft()
    If frmMDI.ActiveForm Is Nothing Then Exit Function
    With frmMDI.ActiveForm
        .rtfText.SelAlignment = rtfLeft 'Set RTB aligment
        ' Set toolbar buttons
        frmMDI.tbrFormat.Buttons("Left").Value = tbrPressed
        frmMDI.tbrFormat.Buttons("Center").Value = tbrUnpressed
        frmMDI.tbrFormat.Buttons("Right").Value = tbrUnpressed
        frmMDI.tbrFormat.Refresh 'Refresh toolbar
        .rtfText.SetFocus 'Set focus
    End With
End Function

Public Function AlignCenter()
    If frmMDI.ActiveForm Is Nothing Then Exit Function
    With frmMDI.ActiveForm
        .rtfText.SelAlignment = rtfCenter 'Set RTB aligment
        ' Set toolbar buttons
        frmMDI.tbrFormat.Buttons("Left").Value = tbrUnpressed
        frmMDI.tbrFormat.Buttons("Center").Value = tbrPressed
        frmMDI.tbrFormat.Buttons("Right").Value = tbrUnpressed
        frmMDI.tbrFormat.Refresh 'Refresh toolbar
        .rtfText.SetFocus 'Set focus
    End With
End Function

Public Function AlignRight()
    If frmMDI.ActiveForm Is Nothing Then Exit Function
    With frmMDI.ActiveForm
        .rtfText.SelAlignment = rtfRight 'Set RTB aligment
        ' Set toolbar buttons
        frmMDI.tbrFormat.Buttons("Left").Value = tbrUnpressed
        frmMDI.tbrFormat.Buttons("Center").Value = tbrUnpressed
        frmMDI.tbrFormat.Buttons("Right").Value = tbrPressed
        frmMDI.tbrFormat.Refresh 'Refresh toolbar
        .rtfText.SetFocus 'Set focus
    End With
End Function

Public Function Bullet()
    If frmMDI.ActiveForm Is Nothing Then Exit Function
    With frmMDI.ActiveForm.rtfText
        'If there is not bullet
        If (IsNull(.SelBullet) = True) Or (.SelBullet = False) Then
            .SelBullet = True 'Put it
            frmMDI.tbrFormat.Buttons("Bullet").Value = tbrPressed
        ElseIf .SelBullet = True Then 'If there is bullet
            .SelBullet = False 'Remove it
            .SelHangingIndent = False
            frmMDI.tbrFormat.Buttons("Bullet").Value = tbrUnpressed
        End If
    End With
End Function
