VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDocument 
   Caption         =   "Document"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8340
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3540
   ScaleWidth      =   8340
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   3540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   6244
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmDocument.frx":014A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmDocument"
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

Public bChanged As Boolean

Private Sub Form_Activate()
    EnableAll 'Enable all menus and toolbars
    SetAll
    GetCurrentLine rtfText
    GetTotalLines rtfText
End Sub

Private Sub Form_Load()
    Form_Resize
    bChanged = False
    ' Get window state
    Me.WindowState = RGGetKeyValue(HKEY_LOCAL_MACHINE, SettingsPath, "Document State", 0)
    EnableAll 'Enable all menus and toolbars
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    rtfText.Width = Me.Width - 120
    rtfText.Height = Me.Height - 380
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If bChanged = True Then 'If document is changed then show prompt
        MsgE "The File " & Me.Caption & " is changed !" & vbCrLf & "Do you want to save changes ?", "ElitePad", 1, False
        ' Get pressed button
        If frmMDI.bYes = True Then SaveDocument: SetAll: DisableAll 'Save document
        If frmMDI.bNo = True Then Unload Me: SetAll: DisableAll 'Close form
        If frmMDI.bCancel = True Then Cancel = True: EnableAll: Exit Sub
    End If
    SetAll
    DisableAll 'Disable all menus and toolbars
    ' Save window position
    If Not Me.WindowState = vbMinimized Then
        RGSetKeyValue HKEY_LOCAL_MACHINE, SettingsPath, "Document State", Me.WindowState
    End If
End Sub

Private Sub rtfText_Change()
    bChanged = True
End Sub

Private Sub rtfText_KeyPress(KeyAscii As Integer)
    SetAll
    GetCurrentLine rtfText
    GetTotalLines rtfText
End Sub

Private Sub rtfText_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then 'If right mouse button is clicked
        PopupMenu frmMDI.mnuPop 'show popup menu
    End If
    SetAll
End Sub

Private Sub rtfText_SelChange()
    SetAll
    GetCurrentLine rtfText
    GetTotalLines rtfText
End Sub
