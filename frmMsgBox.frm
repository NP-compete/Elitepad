VERSION 5.00
Begin VB.Form frmMsgBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ElitePad"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   Icon            =   "frmMsgBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdYes 
      Caption         =   "&Yes"
      Height          =   320
      Left            =   5325
      TabIndex        =   1
      Top             =   147
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   320
      Left            =   5325
      TabIndex        =   3
      Top             =   897
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "&No"
      Height          =   320
      Left            =   5325
      TabIndex        =   2
      Top             =   522
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   320
      Left            =   5325
      TabIndex        =   0
      Top             =   150
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgHelp 
      Height          =   480
      Left            =   83
      Picture         =   "frmMsgBox.frx":000C
      Top             =   147
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'Transparent
      Height          =   975
      Left            =   683
      TabIndex        =   4
      Top             =   147
      Width           =   4575
   End
   Begin VB.Image imgCri 
      Height          =   480
      Left            =   83
      Picture         =   "frmMsgBox.frx":0CD6
      Top             =   147
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmMsgBox"
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

Private Sub cmdYes_Click()
    frmMDI.bYes = True
    Unload Me
End Sub
Private Sub cmdNo_Click()
    frmMDI.bNo = True
    Unload Me
End Sub
Private Sub cmdCancel_Click()
    frmMDI.bCancel = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    With frmMDI
        .bCancel = False
        .bNo = False
        .bYes = False
    End With
    DisableX Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    imgHelp.Visible = False
    imgCri.Visible = False
    lblMsg.Caption = ""
    Me.Caption = ""
    cmdOk.Visible = False
    cmdYes.Visible = False
    cmdNo.Visible = False
    cmdCancel.Visible = False
End Sub
