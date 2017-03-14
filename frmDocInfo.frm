VERSION 5.00
Begin VB.Form frmDocInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Document Properties"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "frmDocInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Dates"
      Height          =   1455
      Left            =   68
      TabIndex        =   19
      Top             =   1935
      Width           =   5175
      Begin VB.Label lblDate 
         Caption         =   "00/00/0000 00:00:00"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   25
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lblDate 
         Caption         =   "00/00/0000 00:00:00"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   24
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label lblDate 
         Caption         =   "00/00/0000 00:00:00"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   23
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Created:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Modified:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Accessed:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   2108
      TabIndex        =   18
      Top             =   5655
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Attributes"
      Height          =   735
      Left            =   68
      TabIndex        =   13
      Top             =   3495
      Width           =   5175
      Begin VB.CheckBox chkAttrib 
         Caption         =   "Archive"
         Height          =   195
         Index           =   4
         Left            =   2640
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chkAttrib 
         Caption         =   "Read Only"
         Height          =   195
         Index           =   3
         Left            =   1440
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chkAttrib 
         Caption         =   "System"
         Height          =   195
         Index           =   2
         Left            =   3840
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chkAttrib 
         Caption         =   "Hidden"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "General"
      Height          =   1095
      Left            =   68
      TabIndex        =   8
      Top             =   735
      Width           =   5175
      Begin VB.Label Label5 
         Caption         =   "Location:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblLocation 
         Height          =   255
         Left            =   1290
         TabIndex        =   11
         Top             =   360
         Width           =   3765
      End
      Begin VB.Label Label6 
         Caption         =   "Size:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblSize 
         Height          =   255
         Left            =   1275
         TabIndex        =   9
         Top             =   675
         Width           =   3885
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   188
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   7
      Top             =   150
      Width           =   615
   End
   Begin VB.Frame Frame4 
      Caption         =   "Document"
      Height          =   1215
      Left            =   68
      TabIndex        =   0
      Top             =   4335
      Width           =   5175
      Begin VB.Label lblWordCount 
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label7 
         Caption         =   "Word Count:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Number of Lines:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblNLines 
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "Characters:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblChr 
         Height          =   255
         Left            =   1560
         TabIndex        =   1
         Top             =   840
         Width           =   2655
      End
   End
   Begin VB.Label Label1 
      Caption         =   "File Name:"
      Height          =   255
      Left            =   1028
      TabIndex        =   27
      Top             =   255
      Width           =   855
   End
   Begin VB.Label lblName 
      Height          =   255
      Left            =   1988
      TabIndex        =   26
      Top             =   255
      Width           =   3135
   End
End
Attribute VB_Name = "frmDocInfo"
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

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FileTime, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const MAX_PATH = 260

Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = DI_MASK Or DI_IMAGE

Private Type FileTime
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
        dwFilechkattrib As Long
        ftCreationTime As FileTime
        ftLastAccessTime As FileTime
        ftLastWriteTime As FileTime
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Public dFileName As String

Private Function FindFile(sFileName As String) As WIN32_FIND_DATA
    Dim Win32Data As WIN32_FIND_DATA
    Dim plngFirstFileHwnd As Long
    Dim plngRtn As Long
    
    ' Find file and get file data
    plngFirstFileHwnd = FindFirstFile(sFileName, Win32Data)
    If plngFirstFileHwnd = 0 Then
        FindFile.cFileName = "Error"
    Else
        FindFile = Win32Data
    End If
    plngRtn = FindClose(plngFirstFileHwnd)
End Function

Private Sub Form_Load()
    dFileName = frmMDI.ActiveForm.Caption 'Get file name
    GetFileInfo
End Sub

Private Function GetFileInfo()
    On Error GoTo FileInfoError
    Dim lSize As Long
    Dim FileTime As SYSTEMTIME
    Dim FileData As WIN32_FIND_DATA
    Dim a() As String
    Dim b() As String
    Dim ChrCount As Long
    Dim WordCount As Long
    Dim I As Long
    
    Screen.MousePointer = vbHourglass 'Set mouse pointer to hourglass
    
    FileData = FindFile(dFileName) 'Find file and get data
    '------GET FILE NAME AND PATH------'
    lblName = GetFTitle(dFileName) 'Get file title
    lblLocation = dFileName 'Get file location
    '------GET FILE SIZE------'
    If FileData.nFileSizeHigh = 0 Then 'Get file size
        lSize = FileData.nFileSizeLow
        lblSize = FormatSize(lSize) 'Format size
    Else
        lSize = FileData.nFileSizeHigh
        lblSize = FormatSize(lSize) 'Format size
    End If
    
    '------GET FILE DATES------'
    ' Created
    FileTimeToSystemTime FileData.ftCreationTime, FileTime
    lblDate(0) = FileTime.wDay & "/" & FileTime.wMonth & "/" & FileTime.wYear & " " & FileTime.wHour & ":" & FileTime.wMinute & ":" & FileTime.wSecond
    ' Modified
    FileTimeToSystemTime FileData.ftLastWriteTime, FileTime
    lblDate(1) = FileTime.wDay & "/" & FileTime.wMonth & "/" & FileTime.wYear & " " & FileTime.wHour & ":" & FileTime.wMinute & ":" & FileTime.wSecond
    ' Accessed
    FileTimeToSystemTime FileData.ftLastAccessTime, FileTime
    lblDate(2) = FileTime.wDay & "/" & FileTime.wMonth & "/" & FileTime.wYear
    
    '------GET FILE ATTRIBUTES------'
    ' Hidden
    If (FileData.dwFilechkattrib And FILE_ATTRIBUTE_HIDDEN) = FILE_ATTRIBUTE_HIDDEN Then chkAttrib(1).Value = 1
    ' System
    If (FileData.dwFilechkattrib And FILE_ATTRIBUTE_SYSTEM) = FILE_ATTRIBUTE_SYSTEM Then chkAttrib(2).Value = 1
    ' Read Only
    If (FileData.dwFilechkattrib And FILE_ATTRIBUTE_READONLY) = FILE_ATTRIBUTE_READONLY Then chkAttrib(3).Value = 1
    ' Archive
    If (FileData.dwFilechkattrib And FILE_ATTRIBUTE_ARCHIVE) = FILE_ATTRIBUTE_ARCHIVE Then chkAttrib(4).Value = 1

    ' Get Word count
    a() = Split(frmMDI.ActiveForm.rtfText.Text, " ") 'Split text to " "
    WordCount = UBound(a)
    For I = 0 To UBound(a)
        If a(I) = "" Then
            WordCount = WordCount - 1
        End If
    Next
    b() = Split(frmMDI.ActiveForm.rtfText.Text, Chr$(10))
    WordCount = WordCount + UBound(b)
    For I = 0 To UBound(b)
        If b(I) = "" Then
            WordCount = WordCount - 1
        End If
    Next
    If WordCount = -2 Then WordCount = -1
    lblWordCount = Format(WordCount + 1, "###,###,###,###")
    ' Get Number of lines
    lblNLines = frmMDI.SB.Panels(5).Text
    ' Get Characters
    ChrCount = SendMessageLong(frmMDI.ActiveForm.rtfText.hWnd, WM_GETTEXTLENGTH, 0, 0)
    lblChr = Format(ChrCount, "###,###,###,###,###")
    '
    Screen.MousePointer = vbDefault 'Set mouse pointer to default
FileInfoError:
    ErrorLog "frmDocInfo\GetFileInfo"
    Screen.MousePointer = vbDefault 'Set mouse pointer to default
End Function

Private Sub chkAttrib_GotFocus(Index As Integer)
    cmdClose.SetFocus 'Don't allow to change attributes
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Picture1_Paint()
    Dim lIcon As Long
    ' Extract assocciate icon from file
    lIcon = ExtractAssociatedIcon(App.hInstance, dFileName, 0&)
    DrawIconEx Picture1.hdc, 0, 0, lIcon, 0, 0, 0, 0, DI_NORMAL 'Draw icon in picturebox
    DestroyIcon lIcon 'Destroy icon
End Sub
