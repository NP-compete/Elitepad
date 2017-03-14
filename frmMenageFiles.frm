VERSION 5.00
Begin VB.Form frmManageFiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manage Files"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "frmMenageFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRename 
      Caption         =   "Rename"
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   1200
      Width           =   6015
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   1560
      Width           =   6015
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdDRename 
      Caption         =   "&Rename ->"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse..."
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdDCopy 
      Caption         =   "&Copy ->"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "From:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "To:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1590
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "File name:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   975
   End
End
Attribute VB_Name = "frmManageFiles"
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

Private Sub cmdCopy_Click()
    CopyFile txtFrom.Text, txtTo.Text 'Copy file
End Sub
Private Sub cmdDelete_Click()
    Me.Height = 1440 'Set height to normal
    cmdCopy.Visible = False 'Hide cmdCopy
    cmdRename.Visible = False 'Hide cmdRename
    DeleteFile txtFileName.Text 'Delete file
End Sub
Private Sub cmdRename_Click()
    RenameFile txtFrom.Text, txtTo.Text 'Rename file
End Sub

Private Sub cmdBrowse_Click()
    frmMDI.CmDlg.DialogTitle = "Select File" 'Set dialog title
    frmMDI.CmDlg.Filter = "All Files (*.*)|*.*" 'Set file filter
    frmMDI.CmDlg.ShowOpen 'Show open dialog
    
    txtFileName = frmMDI.CmDlg.cFileName.Item(1) 'Set selected filename
    txtFrom.Text = txtFileName.Text
End Sub

Private Sub cmdDCopy_Click()
    If cmdDCopy.Caption = "&Copy ->" Then
        Me.Height = 2730 '
        cmdDCopy.Caption = "&Copy <-" 'Change cmdDCopy caption
        cmdDRename.Caption = "&Rename ->" 'Change cmdDRename caption
        cmdCopy.Visible = True 'Show cmdCopy
        cmdRename.Visible = False 'Show cmdRename
    Else
        Me.Height = 1440
        cmdDCopy.Caption = "&Copy ->" 'Change cmdDCopy caption
        cmdDRename.Caption = "&Rename ->" 'Change cmdDRename caption
        cmdCopy.Visible = False 'Show cmdCopy
        cmdRename.Visible = False 'Show cmdRename
    End If
End Sub
Private Sub cmdDRename_Click()
    If cmdDRename.Caption = "&Rename ->" Then
        Me.Height = 2730
        cmdDRename.Caption = "&Rename <-"
        cmdDCopy.Caption = "&Copy ->"
        cmdRename.Visible = True
        cmdCopy.Visible = False
    Else
        Me.Height = 1440
        cmdDCopy.Caption = "&Copy ->"
        cmdDRename.Caption = "&Rename ->"
        cmdRename.Visible = False
        cmdCopy.Visible = False
    End If
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub txtFileName_KeyPress(KeyAscii As Integer)
    txtFrom.Text = txtFileName.Text
End Sub
Private Sub txtFileName_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    txtFrom.Text = txtFileName.Text
End Sub
