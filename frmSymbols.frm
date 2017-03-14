VERSION 5.00
Begin VB.Form frmSymbols 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Symbols"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3180
   Icon            =   "frmSymbols.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   1710
      TabIndex        =   3
      Top             =   3630
      Width           =   1215
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Height          =   375
      Left            =   270
      TabIndex        =   2
      Top             =   3630
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Symbols"
      Height          =   3375
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   2895
      Begin VB.ListBox lstSymbols 
         Height          =   2985
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmSymbols"
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

Private Sub lstSymbols_DblClick()
    ' Insert selected symbol to rtfText
    frmMDI.ActiveForm.rtfText.SelText = Right(lstSymbols.Text, 1)
End Sub

Private Sub cmdInsert_Click()
    ' Insert selected symbol to rtfText
    frmMDI.ActiveForm.rtfText.SelText = Right(lstSymbols.Text, 1)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim I As Integer
    
    'Set font name
    lstSymbols.FontName = frmMDI.ActiveForm.rtfText.Font.Name
    For I = 1 To 255
        ' Fills lstSymbols with Symbols
        If I < 10 Then
            lstSymbols.AddItem I & "     -  " & Chr(I)
        ElseIf I < 100 Then
            lstSymbols.AddItem I & "   -  " & Chr(I)
        Else
            lstSymbols.AddItem I & " -  " & Chr(I)
        End If
    Next I
End Sub
