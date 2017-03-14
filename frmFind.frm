VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBar 
      BorderStyle     =   0  'None
      Height          =   1290
      Left            =   75
      ScaleHeight     =   1290
      ScaleWidth      =   5340
      TabIndex        =   6
      Top             =   900
      Width           =   5340
      Begin VB.CommandButton cmdReplaceAll 
         Caption         =   "Replace &All"
         Height          =   315
         Left            =   4275
         TabIndex        =   13
         Top             =   525
         Width           =   990
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "&Help"
         Height          =   315
         Left            =   4275
         TabIndex        =   12
         Top             =   900
         Width           =   990
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "&Replace..."
         Height          =   315
         Left            =   4275
         TabIndex        =   11
         Top             =   150
         Width           =   990
      End
      Begin VB.Frame Frame1 
         Caption         =   "Search Options"
         Height          =   1215
         Left            =   75
         TabIndex        =   7
         Top             =   0
         Width           =   4065
         Begin VB.CheckBox chkWholeWord 
            Caption         =   "Find Whole Word &Only"
            Height          =   240
            Left            =   150
            TabIndex        =   10
            Top             =   300
            Width           =   1965
         End
         Begin VB.CheckBox chkMatchCase 
            Caption         =   "Match Ca&se"
            Height          =   240
            Left            =   150
            TabIndex        =   9
            Top             =   600
            Width           =   1965
         End
         Begin VB.CheckBox chkNoHighlight 
            Caption         =   "No &Highlight"
            Height          =   240
            Left            =   150
            TabIndex        =   8
            Top             =   900
            Width           =   1965
         End
      End
   End
   Begin VB.ComboBox cboReplace 
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   450
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4350
      TabIndex        =   3
      Top             =   450
      Width           =   990
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   315
      Left            =   4350
      TabIndex        =   1
      Top             =   75
      Width           =   990
   End
   Begin VB.ComboBox cboFind 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   75
      Width           =   3015
   End
   Begin VB.Label lblReplace 
      Caption         =   "Replace &With:"
      Height          =   240
      Left            =   75
      TabIndex        =   4
      Top             =   525
      Width           =   1065
   End
   Begin VB.Label lblFind 
      Caption         =   "Fin&d What:"
      Height          =   240
      Left            =   75
      TabIndex        =   2
      Top             =   150
      Width           =   840
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub cmdFind_Click()
    On Error GoTo FindError
    Dim lngResult As Long
    Dim lngPos As Long
    Dim intOptions As Integer
    ' Set search options
    If chkNoHighlight.Value = 1 Then intOptions = intOptions + 8
    If chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If chkMatchCase.Value = 1 Then intOptions = intOptions + 4

    If cmdFind.Caption = "&Find" Then 'If first time
        ' Get position of the searched word
        lngResult = frmMDI.ActiveForm.rtfText.Find(cboFind.Text, 0, , intOptions)

        If lngResult = -1 Then 'Text not found
            MsgE "Text not found", "ElitePad - Find", 1, True 'Show message
            cmdFind.Caption = "&Find" 'Set caption
            frmMDI.mnuSearchFindNext.Enabled = False 'Disable Find Next menu
        Else 'Text found
            frmMDI.ActiveForm.rtfText.SetFocus 'Set focus to rtfText
            cmdReplace.Enabled = True 'Enable Replace button
            cmdReplaceAll.Enabled = True 'Enable ReplaceAll button
            cmdFind.Caption = "&Find Next" 'Set caption
            frmMDI.mnuSearchFindNext.Enabled = True 'Enable Find Next menu
        End If
    Else 'Find Next
        lngPos = frmMDI.ActiveForm.rtfText.SelStart + frmMDI.ActiveForm.rtfText.SelLength
        lngResult = frmMDI.ActiveForm.rtfText.Find(cboFind.Text, lngPos, , intOptions)

        If lngResult = -1 Then 'Text not found
            MsgE "Text not found", "ElitePad - Find", 1, True 'Show message
            cmdFind.Caption = "&Find" 'Set caption
            cmdReplace.Enabled = False 'Disable Replace button
            cmdReplaceAll.Enabled = False 'Disable ReplaceAll button
            frmMDI.mnuSearchFindNext.Enabled = False 'Disable Find Next menu
        Else 'Text found
            frmMDI.ActiveForm.rtfText.SetFocus 'Set focus to rtfText
            frmMDI.mnuSearchFindNext.Enabled = True 'Enable Find Next menu
        End If
    End If
FindError:
    ErrorLog "frmFind\cmdFind_Click"
End Sub

Private Sub cmdReplace_Click()
    On Error GoTo ReplaceError
    Dim lngResult As Long
    Dim lngPos As Long
    Dim intOptions As Integer
    
    If cmdReplace.Caption = "&Replace..." Then 'Show replace
        cmdReplace.Top = 150 'Set cmdReplace top
        cmdReplace.Caption = "&Replace" 'Set caption
        lblReplace.Visible = True 'Show lblReplace
        cboReplace.Visible = True 'Show cboReplace
        cmdReplaceAll.Visible = True 'Show cmdReplaceAll
        Exit Sub
    End If

    ' Set search options
    If chkNoHighlight.Value = 1 Then intOptions = intOptions + 8
    If chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If chkMatchCase.Value = 1 Then intOptions = intOptions + 4
    
    With frmMDI.ActiveForm
        .rtfText.SelText = cboReplace.Text 'Replace text
        ' Find next
        lngPos = .rtfText.SelStart + .rtfText.SelLength
        ' Get position of the searched word
        lngResult = .rtfText.Find(cboFind.Text, lngPos, , intOptions)

        If lngResult = -1 Then 'Text not found
            MsgE "Text not found", "ElitePad - Replace", 1, True 'Show message
            cmdFind.Caption = "&Find" 'Set caption
            cmdReplace.Enabled = False 'Disable Replace button
            cmdReplaceAll.Enabled = False 'Disable ReplaceAll button
        Else 'Text found
            .rtfText.SetFocus 'Set focus
        End If
    End With
ReplaceError:
    ErrorLog "frmFind\cmdReplace_Click"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdReplaceAll_Click()
    On Error GoTo ReplaceAllError
    Dim intCount As Integer
    Dim lngPos As Long
    Dim intOptions As Integer
    ' Set search options
    If chkNoHighlight.Value = 1 Then intOptions = intOptions + 8
    If chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If chkMatchCase.Value = 1 Then intOptions = intOptions + 4

    intCount = 0
    lngPos = 0
    With frmMDI.ActiveForm
        Do
            If .rtfText.Find(cboFind.Text, lngPos, , intOptions) = -1 Then 'Text not fount
                If intCount > 0 Then 'Show how many replacments have been made
                    MsgE "The specified region has been searched. " & vbCrLf & _
                    intCount & " replacements have been made.", "ElitePad - ReplaceAll", 1, True
                End If
                cmdFind.Caption = "&Find" 'Set caption
                cmdReplace.Enabled = False 'Disable Replace button
                cmdReplaceAll.Enabled = False 'Disable ReplaceAll button
                Exit Do
            Else 'Text found
                lngPos = .rtfText.SelStart + .rtfText.SelLength
                intCount = intCount + 1 'Increase counter by 1
                .rtfText.SelText = cboReplace.Text 'Replace text
            End If
        Loop
    End With
ReplaceAllError:
    ErrorLog "frmFind\cmdReplaceAll_Click"
End Sub

Private Sub Form_Load()
    cmdReplace.Top = 525 'Set cmdReplace top
    lblReplace.Visible = False 'Hide lblReplace
    cboReplace.Visible = False 'Hide cboReplace
    cmdReplaceAll.Visible = False 'Hide cmdReplaceAll
    
    cboFind.AddItem frmMDI.ActiveForm.rtfText.SelText 'Add selected text to find combobox
    cboFind.Text = frmMDI.ActiveForm.rtfText.SelText 'Set text in cbo
End Sub
