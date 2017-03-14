VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4875
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6285
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picAsc 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   225
      ScaleHeight     =   3495
      ScaleWidth      =   5895
      TabIndex        =   35
      Top             =   585
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Frame Frame8 
         Caption         =   "File Association"
         Height          =   3165
         Left            =   90
         TabIndex        =   36
         Top             =   120
         Width           =   5655
         Begin VB.CheckBox chkAscINI 
            Caption         =   "INI Files (INI)"
            Height          =   240
            Left            =   405
            TabIndex        =   42
            Top             =   1665
            Width           =   1995
         End
         Begin VB.CheckBox chkAscBatch 
            Caption         =   "Batch Files (BAT)"
            Height          =   240
            Left            =   405
            TabIndex        =   41
            Top             =   1395
            Width           =   1995
         End
         Begin VB.CheckBox chkAscRich 
            Caption         =   "Rich Text Files (RTF)"
            Height          =   240
            Left            =   405
            TabIndex        =   40
            Top             =   1125
            Width           =   1995
         End
         Begin VB.CheckBox chkAscLog 
            Caption         =   "Log Files (LOG)"
            Height          =   240
            Left            =   405
            TabIndex        =   39
            Top             =   855
            Width           =   1995
         End
         Begin VB.CheckBox chkAscText 
            Caption         =   "Text Files (TXT)"
            Height          =   240
            Left            =   405
            TabIndex        =   38
            Top             =   585
            Width           =   1995
         End
         Begin VB.Label Label5 
            Caption         =   "Select file types you want to associate with ElitePad:"
            Height          =   240
            Left            =   270
            TabIndex        =   37
            Top             =   315
            Width           =   4110
         End
      End
   End
   Begin VB.PictureBox picView 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   195
      ScaleHeight     =   3495
      ScaleWidth      =   5895
      TabIndex        =   11
      Top             =   615
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Frame Frame3 
         Caption         =   "Toolbars"
         Height          =   1575
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   5535
         Begin VB.CheckBox chkWindow 
            Caption         =   "Window"
            Height          =   255
            Left            =   2520
            TabIndex        =   22
            Top             =   1080
            Width           =   2055
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "Edit"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   2055
         End
         Begin VB.CheckBox chkFile 
            Caption         =   "File"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   2055
         End
         Begin VB.CheckBox chkStandard 
            Caption         =   "Standard"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   2055
         End
         Begin VB.CheckBox chkFont 
            Caption         =   "Font"
            Height          =   255
            Left            =   2520
            TabIndex        =   18
            Top             =   720
            Width           =   2055
         End
         Begin VB.CheckBox chkFormat 
            Caption         =   "Format"
            Height          =   255
            Left            =   2520
            TabIndex        =   17
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Other"
         Height          =   1575
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   5535
         Begin VB.CheckBox chkStatusBar 
            Caption         =   "Status Bar"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   2055
         End
         Begin VB.CheckBox chkFileTree 
            Caption         =   "File Tree"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CheckBox chkRuler 
            Caption         =   "Ruler"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1080
            Width           =   2055
         End
      End
   End
   Begin VB.PictureBox picGeneral 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   180
      ScaleHeight     =   3495
      ScaleWidth      =   5895
      TabIndex        =   0
      Top             =   630
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Frame Frame5 
         Caption         =   "Document"
         Height          =   1455
         Left            =   135
         TabIndex        =   5
         Top             =   1980
         Width           =   5655
         Begin VB.ComboBox cmbDocState 
            Height          =   315
            ItemData        =   "frmOptions.frx":000C
            Left            =   360
            List            =   "frmOptions.frx":0019
            TabIndex        =   6
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label3 
            Caption         =   "Select document startup position:"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "General"
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   5655
         Begin VB.CommandButton cmdCleanReg 
            Caption         =   "&Clean Registry"
            Height          =   375
            Left            =   180
            TabIndex        =   4
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CheckBox chkShowTips 
            Caption         =   "Show Tips at Startup"
            Height          =   255
            Left            =   225
            TabIndex        =   3
            Top             =   360
            Width           =   2415
         End
         Begin VB.CheckBox chkStayOnTop 
            Caption         =   "Stay on Top"
            Height          =   255
            Left            =   225
            TabIndex        =   2
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label4 
            Caption         =   "Remove ElitePad key from the registry"
            Height          =   240
            Left            =   1755
            TabIndex        =   34
            Top             =   1170
            Width           =   2760
         End
      End
   End
   Begin VB.PictureBox picEditor 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   195
      ScaleHeight     =   3495
      ScaleWidth      =   5895
      TabIndex        =   23
      Top             =   615
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Frame Frame1 
         Caption         =   "Font"
         Height          =   1455
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   5535
         Begin VB.ComboBox cmbOFontName 
            Height          =   315
            Left            =   1440
            Sorted          =   -1  'True
            TabIndex        =   30
            Top             =   360
            Width           =   3855
         End
         Begin VB.ComboBox cmbOFontSize 
            Height          =   315
            Left            =   1440
            TabIndex        =   29
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Font:"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   400
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Font Size:"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   900
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "View Mode"
         Height          =   1695
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   5535
         Begin VB.OptionButton optNoWrap 
            Caption         =   "No Wrap"
            Height          =   252
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton optWYSIWYG 
            Caption         =   "WYSIWYG"
            Height          =   252
            Left            =   120
            TabIndex        =   26
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton optWordWrap 
            Caption         =   "Word Wrap"
            Height          =   252
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   1695
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   4995
      TabIndex        =   10
      Top             =   4455
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3675
      TabIndex        =   9
      Top             =   4455
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2355
      TabIndex        =   8
      Top             =   4455
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip tbrOptions 
      Height          =   4215
      Left            =   75
      TabIndex        =   33
      Top             =   135
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7435
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      TabFixedWidth   =   1764
      TabMinWidth     =   988
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Editor"
            Key             =   "Editor"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "View"
            Key             =   "View"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Assocciate"
            Key             =   "Assocciate"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
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

Dim VMode As Integer

Private Sub cmdCleanReg_Click()
    RGDeleteKey HKEY_LOCAL_MACHINE, "Software\ElitePad"
End Sub

Private Sub Form_Load()
    ' Center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    ' Set Tabs
    tbrOptions_Click
    GetData ' Get all
End Sub

'------GET ALL------'
Private Function GetData()
    On Error Resume Next
    Dim I As Integer
    ' Get fonts
    For I = 0 To frmMDI.cboFontName.ListCount
        cmbOFontName.AddItem frmMDI.cboFontName.List(I)
    Next I
    ' Get font size
    For I = 0 To frmMDI.cboFontSize.ListCount
        cmbOFontSize.AddItem frmMDI.cboFontSize.List(I)
    Next I
    ' Set font name and size
    cmbOFontName.Text = frmMDI.ActiveForm.rtfText.SelFontName
    cmbOFontSize.Text = frmMDI.ActiveForm.rtfText.SelFontSize
    ' Set view mode
    If frmMDI.mnuViewMode(0).Checked = True Then optNoWrap.Value = True
    If frmMDI.mnuViewMode(1).Checked = True Then optWordWrap.Value = True
    If frmMDI.mnuViewMode(2).Checked = True Then optWYSIWYG.Value = True
    ' Set Tips
    chkShowTips.Value = RGGetKeyValue(HKEY_LOCAL_MACHINE, SettingsPath, "Show Tips at Startup", 1)
    ' Set Toolbars
    If frmMDI.cbrBar.Bands(1).Visible Then chkStandard.Value = 1
    If frmMDI.cbrBar.Bands(2).Visible Then chkFont.Value = 1
    If frmMDI.cbrBar.Bands(3).Visible Then chkFormat.Value = 1
    If frmMDI.cbrBar.Bands(4).Visible Then chkFile.Value = 1
    If frmMDI.cbrBar.Bands(5).Visible Then chkEdit.Value = 1
    If frmMDI.cbrBar.Bands(6).Visible Then chkWindow.Value = 1
    ' Set other
    If frmMDI.SB.Visible Then chkStatusBar.Value = 1
    If frmMDI.cbrRuler.Visible Then chkRuler.Value = 1
    ' Set document state
    cmbDocState.ListIndex = 0
    ' Set stay on top
    chkStayOnTop.Value = RGGetKeyValue(HKEY_LOCAL_MACHINE, ViewPath, "Stay on Top", 0)
End Function

'------SET ALL------'
Private Function SetData()
    On Error Resume Next
    Dim I As Integer
    With frmMDI.ActiveForm
        ' Set font name and size
        .rtfText.Font.Name = cmbOFontName.Text
        .rtfText.Font.Size = cmbOFontSize.Text
        frmMDI.cboFontName.Text = cmbOFontName.Text
        frmMDI.cboFontSize.Text = cmbOFontSize.Text
        frmSymbols.lstSymbols.FontName = cmbOFontName.Text
        RGSetKeyValue HKEY_LOCAL_MACHINE, SettingsPath, "Font Name", cmbOFontName.Text
        RGSetKeyValue HKEY_LOCAL_MACHINE, SettingsPath, "Font Size", cmbOFontSize.Text
        ' Set view mode
        SetViewMode VMode
        For I = 0 To 2
            frmMDI.mnuViewMode(I).Checked = False
        Next
        frmMDI.mnuViewMode(VMode).Checked = True
        RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "ViewMode", Str(VMode)
        ' Set Tips
        RGSetKeyValue HKEY_LOCAL_MACHINE, SettingsPath, "Show Tips at Startup", chkShowTips.Value
    End With
    
    With frmMDI
        '// Set toolbars
        ' Standard
        .cbrBar.Bands(1).Visible = chkStandard.Value
        .mnuViewToolbarStandard.Checked = chkStandard.Value
        RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Standard Toolbar", .mnuViewToolbarStandard.Checked
        ' Font
        .cbrBar.Bands(2).Visible = chkFont.Value
        .mnuViewToolbarFont.Checked = chkFont.Value
        RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Font Toolbar", .mnuViewToolbarFont.Checked
        ' Format
        .cbrBar.Bands(3).Visible = chkFormat.Value
        .mnuViewToolbarFormat.Checked = chkFormat.Value
        RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Format Toolbar", .mnuViewToolbarFormat.Checked
        ' File
        .cbrBar.Bands(4).Visible = chkFile.Value
        .mnuViewToolbarFile.Checked = chkFile.Value
        RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "File Toolbar", .mnuViewToolbarFile.Checked
        ' Edit
        .cbrBar.Bands(5).Visible = chkEdit.Value
        .mnuViewToolbarEdit.Checked = chkEdit.Value
        RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Edit Toolbar", .mnuViewToolbarEdit.Checked
        ' Window
        .cbrBar.Bands(6).Visible = chkWindow.Value
        .mnuViewToolbarWindow.Checked = chkWindow.Value
        RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Window Toolbar", .mnuViewToolbarWindow.Checked
        
        '// Set Other
        ' Status bar
        .SB.Visible = chkStatusBar.Value
        .mnuViewStatusBar.Checked = chkStatusBar.Value
        RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Status bar", .mnuViewStatusBar.Checked
        ' Ruler
        .cbrRuler.Visible = chkRuler.Value
        .mnuViewRuler.Checked = chkRuler.Value
        RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Ruler", .mnuViewRuler.Checked
        ' File Tree
        .picFileBar.Visible = chkFileTree.Value
        .mnuViewFileTree.Checked = chkFileTree.Value
        RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "File Tree", .mnuViewFileTree.Checked
        ' Document State
        RGSetKeyValue HKEY_LOCAL_MACHINE, SettingsPath, "Document State", cmbDocState.ListIndex
        ' Stay on Top
        .mnuViewStayonTop.Checked = chkStayOnTop.Value
        If .mnuViewStayonTop.Checked Then OnTop frmMDI: .mnuViewStayonTop.Checked = True
        If Not .mnuViewStayonTop.Checked Then NotOnTop frmMDI: .mnuViewStayonTop.Checked = False
        RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Stay on Top", chkStayOnTop.Value
        ' Associate Files
        Dim prg As String
        Dim ico As String
        prg = App.Path & "\ElitePad.exe"
        ico = App.Path & "\Doc.ico"
        If chkAscText.Value = 1 Then Associate prg, "txt", "Text Document", ico
        If chkAscRich.Value = 1 Then Associate prg, "rtf", "Rich Text Format", ico
        If chkAscLog.Value = 1 Then Associate prg, "log", "Log File", ico
        If chkAscINI.Value = 1 Then Associate prg, "ini", "Configuration Settings", ico
        If chkAscBatch.Value = 1 Then Associate prg, "bat", "Batch File", ico
    End With
End Function

'------VIEW MODE------'
Private Sub optNoWrap_Click()
    VMode = 0
End Sub
Private Sub optWordWrap_Click()
    VMode = 1
End Sub
Private Sub optWYSIWYG_Click()
    VMode = 2
End Sub

Private Sub cmdOK_Click()
    SetData
    Unload Me
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdApply_Click()
    SetData
End Sub

'------TABS------'
Private Sub tbrOptions_Click()
    ' Change tabs
    If tbrOptions.SelectedItem.Caption = "General" Then
        picGeneral.Visible = True
        picEditor.Visible = False
        picView.Visible = False
        picAsc.Visible = False
    ElseIf tbrOptions.SelectedItem.Caption = "Editor" Then
        picGeneral.Visible = False
        picEditor.Visible = True
        picView.Visible = False
        picAsc.Visible = False
    ElseIf tbrOptions.SelectedItem.Caption = "View" Then
        picGeneral.Visible = False
        picEditor.Visible = False
        picView.Visible = True
        picAsc.Visible = False
    ElseIf tbrOptions.SelectedItem.Caption = "Assocciate" Then
        picGeneral.Visible = False
        picEditor.Visible = False
        picView.Visible = False
        picAsc.Visible = True
    End If
End Sub
