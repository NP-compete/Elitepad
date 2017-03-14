VERSION 5.00
Begin VB.Form frmTimeDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Time and Date"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2940
   Icon            =   "frmTimeDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   2940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   1590
      TabIndex        =   8
      Top             =   3390
      Width           =   1215
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Default         =   -1  'True
      Height          =   375
      Left            =   150
      TabIndex        =   7
      Top             =   3390
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Formats:"
      Height          =   3135
      Left            =   150
      TabIndex        =   5
      Top             =   150
      Width           =   2655
      Begin VB.ListBox lstDates 
         Height          =   2790
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Custom"
      Height          =   2055
      Left            =   150
      TabIndex        =   0
      Top             =   3870
      Width           =   2655
      Begin VB.TextBox txtCustom 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton cmdInsCusFor 
         Caption         =   "Insert Custom Format"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label lblSample 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Enter custom format:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmTimeDate"
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

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdInsCusFor_Click()
    frmMDI.ActiveForm.rtfText.SelText = lblSample.Caption 'Insert custom format
    Unload Me
End Sub

Private Sub cmdInsert_Click()
    frmMDI.ActiveForm.rtfText.SelText = lstDates.Text 'Insert time and date
    Unload Me
End Sub

Private Sub lstDates_DblClick()
    frmMDI.ActiveForm.rtfText.SelText = lstDates.Text 'Insert time and date
End Sub

Private Sub Form_Load()
    'Fill listbox with time and date formats
    With lstDates
        .AddItem Format(Date, "dd/mm/yy"), 0
        .AddItem Format(Date, "dd mmmm yyyy"), 1
        .AddItem Format(Date, "dd mmmm, yyyy"), 2
        .AddItem Format(Date, "mmmm dd, yyyy"), 3
        .AddItem Format(Date, "dd/mmm/yy"), 4
        .AddItem Format(Date, "mmmm, yy"), 5
        .AddItem Format(Date, "mmm/yy"), 6
        .AddItem Format(Date, "dd/mm/yy") & " " & Format(Time, "hh:mm"), 7
        .AddItem Format(Date, "dd/mm/yy") & " " & Format(Time, "hh:mm:ss"), 8
        .AddItem Format(Time, "hh:mm"), 9
        .AddItem Format(Time, "hh:mm:ss"), 10
        .AddItem Format(Time, "hh:mm AMPM"), 11
        .AddItem Format(Time, "hh:mm:ss AMPM"), 12
    End With
End Sub

Private Sub txtCustom_KeyPress(KeyAscii As Integer)
    lblSample.Caption = Format(Now, txtCustom.Text) 'Update lblSample with custom format
End Sub

Private Sub txtCustom_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblSample.Caption = Format(Now, txtCustom.Text) 'Update lblSample with custom format
End Sub
