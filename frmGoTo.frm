VERSION 5.00
Begin VB.Form frmGoTo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Go To..."
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   Icon            =   "frmGoTo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3075
      TabIndex        =   6
      Top             =   900
      Width           =   1095
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&GO"
      Default         =   -1  'True
      Height          =   375
      Left            =   1755
      TabIndex        =   5
      Top             =   900
      Width           =   1095
   End
   Begin VB.TextBox txtGo 
      Height          =   285
      Left            =   1755
      TabIndex        =   4
      Top             =   270
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Go to what:"
      Height          =   1095
      Left            =   75
      TabIndex        =   0
      Top             =   150
      Width           =   1455
      Begin VB.OptionButton GotoStart 
         Caption         =   "&Start"
         Height          =   252
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   720
      End
      Begin VB.OptionButton GotoLine 
         Caption         =   "&Line"
         Height          =   252
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton GoToEnd 
         Caption         =   "End"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Enter line number you want to go to"
      Height          =   255
      Left            =   1635
      TabIndex        =   7
      Top             =   630
      Width           =   2655
   End
End
Attribute VB_Name = "frmGoTo"
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

Private Sub cmdGo_Click()
    On Error GoTo GoToError
    Dim lngStart As Long

    With frmMDI.ActiveForm.rtfText
        If GotoLine Then 'If Go To line is checked
            'Get pos of start of the line
            lngStart = SendMessage(.hWnd, EM_LINEINDEX, txtGo.Text - 1, 0&)
            If lngStart = -1 Then 'Invalid line number
                MsgE "Invalid line number!", "ElitePad - Go To", 0, True
                Exit Sub
            End If
            .SelStart = lngStart 'Go To line
        ElseIf GotoStart Then 'Go To start of the document
            .SelStart = 0
        ElseIf GoToEnd Then 'Go To end of the document
            .SelStart = Len(.Text)
        End If
        .SetFocus 'Set focus
        Unload Me
    End With
GoToError:
    If Err.Number = 13 Then
        MsgE "You can only type numbers!!!", "ElitePad - Go To", 1, True
        Exit Sub
    End If
End Sub

Private Sub txtGo_Change()
    On Error Resume Next
    If Not txtGo.Text = "" Or txtGo.Enabled = False Then
         cmdGo.Enabled = True 'Enable cmdGo button
    ElseIf txtGo.Text = "" Then
         cmdGo.Enabled = False 'Disable cmdGo button
    End If
End Sub

Private Sub GotoLine_Click()
    txtGo.Enabled = True 'Enable txtGo
    txtGo_Change
End Sub

Private Sub GotoStart_Click()
    txtGo.Enabled = False 'Disable txtGo
    txtGo_Change
End Sub

Private Sub GoToEnd_Click()
    txtGo.Enabled = False 'Disable txtGo
    txtGo_Change
End Sub

Private Sub Form_Load()
    GotoLine_Click
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
