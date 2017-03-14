VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "ElitePad"
   ClientHeight    =   7080
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9945
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picFileBar 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   4890
      Left            =   0
      ScaleHeight     =   4890
      ScaleWidth      =   2355
      TabIndex        =   12
      Top             =   1920
      Width           =   2355
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMDI.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMDI.frx":2ABE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMDI.frx":5272
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMDI.frx":7A26
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMDI.frx":A1DA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.FileListBox File1 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Hidden          =   -1  'True
         Left            =   0
         System          =   -1  'True
         TabIndex        =   14
         Top             =   3375
         Width           =   2310
      End
      Begin MSComctlLib.TreeView tvwFolders 
         CausesValidation=   0   'False
         Height          =   3300
         Left            =   0
         TabIndex        =   13
         Top             =   75
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   5821
         _Version        =   393217
         Indentation     =   0
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.DirListBox Dir1 
         Height          =   315
         Left            =   0
         TabIndex        =   16
         Top             =   375
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin ElitePad.epCmDlg CmDlg 
      Left            =   6150
      Top             =   3675
      _ExtentX        =   661
      _ExtentY        =   635
   End
   Begin ComCtl3.CoolBar cbrBar 
      Align           =   1  'Align Top
      Height          =   1500
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   2646
      BandCount       =   6
      _CBWidth        =   9945
      _CBHeight       =   1500
      _Version        =   "6.0.8169"
      Child1          =   "tbrStandard"
      MinHeight1      =   330
      Width1          =   30
      NewRow1         =   0   'False
      Child2          =   "picFont"
      MinHeight2      =   360
      Width2          =   3300
      NewRow2         =   -1  'True
      Child3          =   "tbrFormat"
      MinHeight3      =   330
      Width3          =   2175
      NewRow3         =   0   'False
      Child4          =   "tbrFile"
      MinHeight4      =   330
      Width4          =   2850
      NewRow4         =   -1  'True
      Visible4        =   0   'False
      Child5          =   "tbrEdit"
      MinHeight5      =   330
      Width5          =   5085
      NewRow5         =   0   'False
      Visible5        =   0   'False
      Child6          =   "tbrWindow"
      MinHeight6      =   330
      Width6          =   2205
      NewRow6         =   -1  'True
      Visible6        =   0   'False
      Begin VB.PictureBox picFont 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   165
         ScaleHeight     =   360
         ScaleWidth      =   3105
         TabIndex        =   8
         Top             =   390
         Width           =   3105
         Begin VB.ComboBox cboFontSize 
            Height          =   315
            Left            =   2250
            TabIndex        =   10
            Top             =   23
            Width           =   690
         End
         Begin VB.ComboBox cboFontName 
            Height          =   315
            Left            =   75
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   23
            Width           =   2115
         End
      End
      Begin MSComctlLib.Toolbar tbrWindow 
         Height          =   330
         Left            =   165
         TabIndex        =   7
         Top             =   1140
         Width           =   9690
         _ExtentX        =   17092
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imgWindow"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "THorizontally"
               Object.ToolTipText     =   "Tile windows horizontally"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "TVertically"
               Object.ToolTipText     =   "Tile windows vertically"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cascade"
               Object.ToolTipText     =   "Cascade windows"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Minimize"
               Object.ToolTipText     =   "Minimize windows"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Restore"
               Object.ToolTipText     =   "Restore windows"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Close"
               Object.ToolTipText     =   "Close the active window"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbrFile 
         Height          =   330
         Left            =   165
         TabIndex        =   6
         Top             =   780
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imgFile"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "New"
               Object.ToolTipText     =   "Create a new document"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               Object.ToolTipText     =   "Open an existing document"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Close"
               Object.ToolTipText     =   "Close the active document"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               Object.ToolTipText     =   "Save the active document"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SaveAll"
               Object.ToolTipText     =   "Save all documents"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "MenageFiles"
               Object.ToolTipText     =   "Copy, delete or rename files"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Print"
               Object.ToolTipText     =   "Print the active document"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbrEdit 
         Height          =   330
         Left            =   3045
         TabIndex        =   5
         Top             =   780
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imgStandard"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cut"
               Object.ToolTipText     =   "Cut the selection to the Clipboard"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Copy"
               Object.ToolTipText     =   "Copy the selection to the Clipboard"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Paste"
               Object.ToolTipText     =   "Insert Clipboard contents"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Undo"
               Object.ToolTipText     =   "Undo the last action"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Redo"
               Object.ToolTipText     =   "Redo the previously undone action"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Outdent"
               Object.ToolTipText     =   "Reduce indentation of selected lines"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Indent"
               Object.ToolTipText     =   "Increase indentation of selected lines"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbrFormat 
         Height          =   330
         Left            =   3495
         TabIndex        =   4
         Top             =   405
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imgFormat"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Find"
               Object.ToolTipText     =   "Find the specified text"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bold"
               Object.ToolTipText     =   "Make selected text bold"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Italic"
               Object.ToolTipText     =   "Make selected text italic"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Underline"
               Object.ToolTipText     =   "Make selected text underline"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "StrikeThru"
               Object.ToolTipText     =   "Make selected text strike through"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Left"
               Object.ToolTipText     =   "Align selected lines to left"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Center"
               Object.ToolTipText     =   "Align selected lines to center"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Right"
               Object.ToolTipText     =   "Align selected lines to right"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bullet"
               Object.ToolTipText     =   "Set bullets"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbrStandard 
         Height          =   330
         Left            =   165
         TabIndex        =   3
         Top             =   30
         Width           =   9690
         _ExtentX        =   17092
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imgStandard"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   21
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "New"
               Object.ToolTipText     =   "Create a new document"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               Object.ToolTipText     =   "Open an existing document"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               Object.ToolTipText     =   "Save the active document"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Print"
               Object.ToolTipText     =   "Print the active document"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FullScreen"
               Object.ToolTipText     =   "View the active document in full screen mode"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "WordWrap"
               Object.ToolTipText     =   "Toggle Word Wrap"
               ImageIndex      =   6
               Style           =   1
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cut"
               Object.ToolTipText     =   "Cut the selection to the Clipboard"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Copy"
               Object.ToolTipText     =   "Copy the selection to the Clipboard"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Paste"
               Object.ToolTipText     =   "Insert Clipboard contents"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Undo"
               Object.ToolTipText     =   "Undo the last action"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Redo"
               Object.ToolTipText     =   "Redo the previously undone action"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Outdent"
               Object.ToolTipText     =   "Reduce indentation of selected lines"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Indent"
               Object.ToolTipText     =   "Increase indentation of selected lines"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Help"
               Object.ToolTipText     =   "Display help contents"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgFormat 
      Left            =   7575
      Top             =   3525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":C98E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":CAEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":CC46
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":CDA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":CEFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":D01E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":D17A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":D2D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":D432
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgWindow 
      Left            =   7275
      Top             =   3525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":D58E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":D6EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":D846
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":D9AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":DB0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":DC66
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgFile 
      Left            =   6975
      Top             =   3525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":DDC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":DF1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":E07A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":E616
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":E772
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":ED0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":F1EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgStandard 
      Left            =   6675
      Top             =   3525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":F346
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":F4A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":F5FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":F75A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":F8B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":FA12
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":FDAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":FF0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":10066
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":101C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1031E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1047A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":10656
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":10832
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrRuler 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   1500
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   741
      BandCount       =   1
      _CBWidth        =   9945
      _CBHeight       =   420
      _Version        =   "6.0.8169"
      MinHeight1      =   360
      Width1          =   1440
      NewRow1         =   0   'False
      Begin VB.PictureBox picInsert 
         Height          =   255
         Left            =   1125
         ScaleHeight     =   195
         ScaleWidth      =   390
         TabIndex        =   17
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin RichTextLib.RichTextBox rtfTemp 
         CausesValidation=   0   'False
         Height          =   240
         Left            =   825
         TabIndex        =   11
         Top             =   75
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   423
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmMDI.frx":1098E
      End
      Begin VB.PictureBox picRuler 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   120
         Picture         =   "frmMDI.frx":10A57
         ScaleHeight     =   270
         ScaleWidth      =   11490
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   80
         Width           =   11490
      End
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   18
      Top             =   6810
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "For Help, press F1"
            TextSave        =   "For Help, press F1"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Line #:"
            TextSave        =   "Line #:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1587
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "Total lines:"
            TextSave        =   "Total lines:"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1587
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1766
            TextSave        =   "12:48 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileCloseAll 
         Caption         =   "Clos&e All"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Save A&ll"
         Shortcut        =   +^{F12}
      End
      Begin VB.Menu mnuFileSaveSelectionAs 
         Caption         =   "Sa&ve Selection As..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRevert 
         Caption         =   "Revert to Save&d"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileManageFiles 
         Caption         =   "Mana&ge Files..."
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintConfig 
         Caption         =   "Print Configura&tion"
         Begin VB.Menu mnuFilePageSetup 
            Caption         =   "Page Set&up..."
         End
         Begin VB.Menu mnuFilePrintSetup 
            Caption         =   "Print Se&tup..."
         End
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   "Recent F&iles"
         Begin VB.Menu mnuFileMRUTemp 
            Caption         =   "--Recent Files--"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuFileMRUItem 
            Caption         =   ""
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFileMRUItem 
            Caption         =   ""
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFileMRUItem 
            Caption         =   ""
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFileMRUItem 
            Caption         =   ""
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFileMRUItem 
            Caption         =   ""
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFileMRUItem 
            Caption         =   ""
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFileMRUItem 
            Caption         =   ""
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFileMRUItem 
            Caption         =   ""
            Index           =   8
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditMarkClean 
         Caption         =   "Mar&k Clean"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSearchFindNext 
         Caption         =   "Find &Next"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSearchBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchReplace 
         Caption         =   "&Replace..."
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSearchBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchGoTo 
         Caption         =   "&Go To..."
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "&Status Bar"
      End
      Begin VB.Menu mnuViewRuler 
         Caption         =   "&Ruler"
      End
      Begin VB.Menu mnuViewFileTree 
         Caption         =   "Fi&le Tree View"
      End
      Begin VB.Menu mnuViewFullScreen 
         Caption         =   "&Full Screen"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewToolbars 
         Caption         =   "T&oolbars"
         Begin VB.Menu mnuViewToolbarStandard 
            Caption         =   "&Standard"
         End
         Begin VB.Menu mnuViewToolbarFile 
            Caption         =   "&File"
         End
         Begin VB.Menu mnuViewToolbarEdit 
            Caption         =   "&Edit"
         End
         Begin VB.Menu mnuViewToolbarFormat 
            Caption         =   "Fo&rmat"
         End
         Begin VB.Menu mnuViewToolbarFont 
            Caption         =   "F&ont"
         End
         Begin VB.Menu mnuViewToolbarWindow 
            Caption         =   "&Window"
         End
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStayonTop 
         Caption         =   "Stay on &Top"
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "&No Wrap"
         Index           =   0
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "&Word Wrap"
         Index           =   1
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "WY&SIWYG"
         Index           =   2
      End
      Begin VB.Menu mnuViewBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewDocumentProperties 
         Caption         =   "Document P&roperties"
         Shortcut        =   +{F3}
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuInsertTimeDate 
         Caption         =   "Time and &Date..."
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuInsertBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertPicture 
         Caption         =   "&Picture..."
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnuInsertBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertTextFile 
         Caption         =   "&Text File..."
      End
      Begin VB.Menu mnuInsertPathandFile 
         Caption         =   "Path and &File"
      End
      Begin VB.Menu mnuInsertBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertSymbols 
         Caption         =   "&Symbols..."
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "F&ormat"
      Begin VB.Menu mnuFormatFont 
         Caption         =   "&Font..."
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuFormatBullet 
         Caption         =   "&Bullet"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFormatBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatUpper 
         Caption         =   "To &Upper Case"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuFormatLower 
         Caption         =   "To &Lower Case"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuFormatBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatScript 
         Caption         =   "&Script"
         Begin VB.Menu mnuFormatScriptNoScript 
            Caption         =   "&No Scripting"
         End
         Begin VB.Menu mnuFormatScriptSuperScript 
            Caption         =   "&SuperScript"
         End
         Begin VB.Menu mnuFormatScriptSubScript 
            Caption         =   "S&ubScript"
         End
      End
      Begin VB.Menu mnuFormatBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatIndent 
         Caption         =   "Increase &Indent"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuFormatOutdent 
         Caption         =   "&Reduce Indent"
         Shortcut        =   +{F1}
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Window"
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
      Begin VB.Menu mnuWindowBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowMinimizeAll 
         Caption         =   "&Minimize All"
      End
      Begin VB.Menu mnuWindowRestoreAll 
         Caption         =   "&Restore All"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "&Index"
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Search"
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpTipoftheDay 
         Caption         =   "Tip of the &Day"
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About ElitePad..."
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuPopBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuPopCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPopPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuPopBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuPopSelectAll 
         Caption         =   "&Select All"
      End
      Begin VB.Menu mnuPopBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopCase 
         Caption         =   "C&hange Case"
         Begin VB.Menu mnuPopCaseUpper 
            Caption         =   "&Upper Case"
         End
         Begin VB.Menu mnuPopCaseLower 
            Caption         =   "&Lower Case"
         End
      End
   End
End
Attribute VB_Name = "frmMDI"
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
' For the custom MsgBox
Public bYes, bNo, bCancel As Boolean
Public MRUList As Collection

Private Sub cboFontName_Click()
    ActiveForm.rtfText.SelFontName = cboFontName.Text 'Set selected font name
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, SettingsPath, "Font Name", cboFontName.Text
End Sub

Private Sub cboFontSize_Click()
    ActiveForm.rtfText.SelFontSize = cboFontSize.Text 'Set selected font size
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, SettingsPath, "Font Size", cboFontSize.Text
End Sub

Private Sub File1_Click()
    Dim fType As String
    Dim strFile As String
    
    ' Get file name
    If Right(File1.Path, 1) = "\" Then
        strFile = File1.Path & File1.FileName
    Else
        strFile = File1.Path & "\" & File1.FileName
    End If
    
    ' Get file extension
    If UCase(Right(strFile, 3)) = "RTF" Then
        fType = rtfText
    Else
        fType = rtfText
    End If
    CreateNewDocument
    ActiveForm.rtfText.LoadFile strFile, fType 'Load file
    ActiveForm.Caption = strFile 'Set form caption
    ActiveForm.bChanged = False 'Set bChanged flag to false
End Sub

Private Sub MDIForm_Load()
    RGCreateKey HKEY_LOCAL_MACHINE, MRUPath 'Create MRU key
    RGCreateKey HKEY_LOCAL_MACHINE, ViewPath 'Create View key
    RGCreateKey HKEY_LOCAL_MACHINE, SettingsPath 'Create Settings key
    CreateNewDocument
    DisableAll 'Disable All menus and toolbars
    GetDrives 'Load drives into FileBar
    ' Check command line
    If Command$ <> "" Then CommLineFile
    ' Show tip of the day
    If Not RGGetKeyValue(HKEY_LOCAL_MACHINE, SettingsPath, "Show Tips at Startup", 1) = 0 Then
        frmTip.Show 1
    End If
    gHW = Me.hWnd 'Store a handle to this form
    Hook 'Call this Sub procedure to begin hooking into messages
End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim OLEFilename As String
    Dim fType As String
    Dim I As Integer
    
    For I = 1 To Data.Files.Count
        If Data.GetFormat(vbCFFiles) Then
            OLEFilename = Data.Files(I)
        End If
    
        'Get file extension
        Select Case UCase(Right(OLEFilename, 3))
            Case "RTF"
                fType = rtfRTF
            Case Else
                fType = rtfText
        End Select
        
        On Error GoTo errexit
        CreateNewDocument
        ActiveForm.rtfText.LoadFile OLEFilename, fType 'Load file
        ActiveForm.Caption = OLEFilename 'Set caption
        ActiveForm.bChanged = False 'Set bChanged flag to false
    Next I
errexit:
    Exit Sub
End Sub

Private Sub MDIForm_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    If Not Data.GetFormat(vbCFFiles) Then Effect = vbDropEffectNone
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim Form As Form
    For Each Form In Forms
        If Form.Name <> Me.Name Then
            Unload Form
            Set Form = Nothing
        End If
    Next Form
    Unhook 'Call this Sub procedure to cease hooking into messages
End Sub

Public Function CommLineFile()
    Dim sFile As String
    Dim fType As String
    
    If Command$ <> "" Then
        sFile = Command$ 'Get command line filename
    End If
  
    'Get file extension
    Select Case UCase(Right(sFile, 3))
        Case "RTF"
            fType = rtfRTF
        Case Else
            fType = rtfText
    End Select
    
    On Error GoTo CommResume
    ActiveForm.rtfText.LoadFile sFile, fType 'Load file
    ActiveForm.Caption = sFile 'Set caption
    ActiveForm.bChanged = False 'Set bChanged flag to false
CommResume:
    If Left(sFile, 1) = Chr(34) Then 'Remove "" from filename
        sFile = Right(sFile, Len(sFile) - 1)
        sFile = Left(sFile, Len(sFile) - 1)
    End If
    ActiveForm.rtfText.LoadFile sFile, fType 'Load file
    ActiveForm.Caption = sFile 'Set caption
    ActiveForm.bChanged = False 'Set bChanged flag to false
End Function

Private Sub GetDrives()
    Dim I As Integer
    Dim strPath As String
    Dim IconN As Integer
    tvwFolders.Nodes.Clear
    
    For I = 0 To Drive1.ListCount - 1 'Get all drives
        ' Get drive
        strPath = UCase(Left(Drive1.List(I), 1)) & ":\"

        Select Case GetDriveType(strPath) 'Check to the type of the drive
            Case 2 'Diskette drive
                IconN = 1
            Case 3 'Hard Disk
                IconN = 2
            Case 5 'CDROM drive
                IconN = 3
            Case Else
                IconN = 2
        End Select
        
        ' Add drive
        tvwFolders.Nodes.Add , , strPath, UCase(Drive1.List(I)), IconN
        tvwFolders.Nodes.Add strPath, tvwChild, ""
    Next
End Sub

Private Sub picFileBar_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    ' Resize file filebar
    File1.Height = picFileBar.Height - tvwFolders.Height - 100
End Sub

'//////  TOOLBARS //////'
'// STANDARD
Private Sub tbrStandard_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "New"
        mnuFileNew_Click
    Case "Open"
        mnuFileOpen_Click
    Case "Save"
        mnuFileSave_Click
    Case "Print"
        mnuFilePrint_Click
    Case "FullScreen"
        mnuViewFullScreen_Click
    Case "WordWrap"
        If tbrStandard.Buttons("WordWrap").Value = tbrPressed Then
            SetViewMode 1
            tbrStandard.Buttons("WordWrap").Value = tbrPressed
        Else
            SetViewMode 0
            tbrStandard.Buttons("WordWrap").Value = tbrUnpressed
        End If
    Case "Cut"
        mnuEditCut_Click
    Case "Copy"
        mnuEditCopy_Click
    Case "Paste"
        mnuEditPaste_Click
    Case "Undo"
        mnuEditUndo_Click
    Case "Redo"
        mnuEditRedo_Click
    Case "Outdent"
        mnuFormatOutdent_Click
    Case "Indent"
        mnuFormatIndent_Click
    Case "Help"
        mnuHelpContents_Click
    End Select
End Sub
'// FORMAT
Private Sub tbrFormat_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Find"
        mnuSearchFind_Click
    Case "Bold"
        Bold
    Case "Italic"
        Italic
    Case "Underline"
        Underline
    Case "StrikeThru"
        Strikethru
    Case "Left"
        AlignLeft
    Case "Center"
        AlignCenter
    Case "Right"
        AlignRight
    Case "Bullet"
        mnuFormatBullet_Click
    End Select
End Sub
'// FILE
Private Sub tbrFile_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "New"
        mnuFileNew_Click
    Case "Open"
        mnuFileOpen_Click
    Case "Close"
        mnuFileClose_Click
    Case "Save"
        mnuFileSave_Click
    Case "SaveAll"
        mnuFileSaveAs_Click
    Case "MenageFiles"
        mnuFileManageFiles_Click
    Case "Print"
        mnuFilePrint_Click
    End Select
End Sub
'// EDIT
Private Sub tbrEdit_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Cut"
        mnuEditCut_Click
    Case "Copy"
        mnuEditCopy_Click
    Case "Paste"
        mnuEditPaste_Click
    Case "Undo"
        mnuEditUndo_Click
    Case "Redo"
        mnuEditRedo_Click
    Case "Outdent"
        mnuFormatOutdent_Click
    Case "Indent"
        mnuFormatIndent_Click
    End Select
End Sub
'// WINDOW
Private Sub tbrWindow_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Cascade"
        mnuWindowCascade_Click
    Case "THorizontally"
        mnuWindowTileHorizontal_Click
    Case "TVertically"
        mnuWindowTileVertical_Click
    Case "Minimize"
        mnuWindowMinimizeAll_Click
    Case "Restore"
        mnuWindowRestoreAll_Click
    Case "Close"
        mnuFileClose_Click
    End Select
End Sub

Private Sub tvwFolders_Expand(ByVal Node As MSComctlLib.Node)
    On Error GoTo TVError
    Dim I As Integer
    Dim strRelative As String
    Dim strFolderName As String
    Dim intFolderPos As Integer
    Dim strNewPath As String

    MousePointer = vbHourglass 'Change mouse pointer to hourglass

    ' Add folders
    If Node.Child.Text = "" Then
             
        tvwFolders.Nodes.Remove Node.Child.Index
        strRelative = Node.Key
        Dir1.Path = strRelative
        intFolderPos = Len(strRelative) + 1

        For I = 0 To Dir1.ListCount - 1
            strFolderName = Mid(Dir1.List(I), intFolderPos)
            strNewPath = strRelative & strFolderName & "\"
            tvwFolders.Nodes.Add strRelative, tvwChild, strNewPath, strFolderName, 4
            Dir1.Path = strNewPath

            If Dir1.ListCount > 0 Then
                tvwFolders.Nodes.Add strNewPath, tvwChild, , ""
                tvwFolders.Nodes(strNewPath).ExpandedImage = 5
            End If

            Dir1.Path = strRelative
        Next

    End If
    MousePointer = vbDefault 'Change mouse pointer to default
    Exit Sub
TVError:
    If Err.Number = 68 Then MsgBox "Device not avaible!", vbCritical, "ElitePad v1.3"
    MousePointer = vbDefault 'Change mouse pointer to hourglass
End Sub

Private Sub tvwFolders_NodeClick(ByVal Node As MSComctlLib.Node)
    File1.Path = Node.Key 'Set path
End Sub

'// FILE MENU
Private Sub mnuFileNew_Click()
    CreateNewDocument 'Call CreateNewDocument function
End Sub
Private Sub mnuFileOpen_Click()
    OpenDocument 'Call OpenDocument function
End Sub
Private Sub mnuFileClose_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    ' Unload the active document
    Unload ActiveForm
End Sub
Private Sub mnuFileCloseAll_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    ' Unload all open documents
    While Forms.Count > 1
        Unload ActiveForm
    Wend
End Sub
Private Sub mnuFileSave_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    SaveDocument 'Call SaveDocument function
End Sub
Private Sub mnuFileSaveAs_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    SaveDocumentAs 'Call SaveDocumentAs function
End Sub
Private Sub mnuFileSaveAll_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    SaveAllDocuments 'Call SaveAllDocuments function
End Sub
Private Sub mnuFileSaveSelectionAs_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    SaveSelectionAs 'Call SaveSelectionAs function
End Sub
Private Sub mnuFileRevert_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    RevertDocument 'Call RevertDocument function
End Sub
Private Sub mnuFileManageFiles_Click()
    frmManageFiles.Show , Me 'Show Manage Files form
End Sub
Private Sub mnuFilePageSetup_Click()
    CmDlg.ShowPageSetup 'Show Page Setup dialog
End Sub
Private Sub mnuFilePrintSetup_Click()
    CmDlg.ShowPrinter 'Show Printer dialog
End Sub
Private Sub mnuFilePrint_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    PrintRTF ActiveForm.rtfText, 720, 720, 720, 720 'Call PrintRTF sub
End Sub
Private Sub mnuFileMRUItem_Click(Index As Integer)
    On Error GoTo MRUError
    Dim fType As String
    
    If FileExists(mnuFileMRUItem(Index).Tag) = True Then 'If file exists
        ' Get file extension
        If UCase(Right(mnuFileMRUItem(Index).Tag, 3)) = "RTF" Then
            fType = rtfText
        Else
            fType = rtfText
        End If

        CreateNewDocument 'Create a new document
        ActiveForm.rtfText.LoadFile mnuFileMRUItem(Index).Tag, fType 'Load file
        ActiveForm.Caption = mnuFileMRUItem(Index).Tag 'Set caption
        ActiveForm.bChanged = False 'Set bChanged flag to false
    Else 'If file doesn't exists show the message
        MsgE "File doesn't exists!", "ElitePad - MRUItem", 0, True
    End If
MRUError:
    ErrorLog "frmMDI\mnuFileMRUItem_Click"
End Sub
Private Sub mnuFileExit_Click()
    Unload Me
End Sub

'// EDIT MENU
Private Sub mnuEditUndo_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    SendMessage ActiveForm.rtfText.hWnd, EM_UNDO, 0, 0&
End Sub
Private Sub mnuEditRedo_Click()
    MsgE "This function has not been implanted yet!", "ElitePad", 1, True
End Sub
Private Sub mnuEditCut_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    SendMessage ActiveForm.rtfText.hWnd, WM_CUT, 0&, 0& 'Cut
End Sub
Private Sub mnuEditCopy_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    SendMessage ActiveForm.rtfText.hWnd, WM_COPY, 0&, 0& 'Copy
End Sub
Private Sub mnuEditPaste_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    SendMessage ActiveForm.rtfText.hWnd, WM_PASTE, 0&, 0& 'Paste
End Sub
Private Sub mnuEditDelete_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    SendMessage ActiveForm.rtfText.hWnd, WM_CLEAR, 0&, 0& 'Delete
End Sub
Private Sub mnuEditSelectAll_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    ActiveForm.rtfText.SelStart = 0 'Set the start pos of the selection
    ActiveForm.rtfText.SelLength = Len(ActiveForm.rtfText) 'Set length of the selection
End Sub
Private Sub mnuEditMarkClean_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    ActiveForm.bChanged = False 'Set bChanged flag to False
    mnuEditMarkClean.Enabled = False 'Disable menu
End Sub

'// SEARCH MENU
Private Sub mnuSearchFind_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    frmFind.Show , Me
End Sub
Private Sub mnuSearchFindNext_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    On Error GoTo FindNextError
    Dim lngResult As Integer
    Dim lngPos As Integer
    Dim intOptions As Integer
    ' Set search options
    If frmFind.chkNoHighlight.Value = 1 Then intOptions = intOptions + 8
    If frmFind.chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If frmFind.chkMatchCase.Value = 1 Then intOptions = intOptions + 4

    lngPos = ActiveForm.rtfText.SelStart + ActiveForm.rtfText.SelLength
    ' Get position of the searched word
    lngResult = ActiveForm.rtfText.Find(frmFind.cboFind.Text, lngPos, , intOptions)

    If lngResult = -1 Then 'Text not found
        MsgE "Text not found", "ElitePad - FindNext", 1, True 'Show message
        frmFind.cmdFind.Caption = "&Find" 'Set caption
        frmFind.cmdReplace.Enabled = False 'Disable Replace button
        frmFind.cmdReplaceAll.Enabled = False 'Disable ReplaceAll button
        mnuSearchFindNext.Enabled = False 'Disable Find Next menu
    Else
        ActiveForm.rtfText.SetFocus 'Set focus
    End If
FindNextError:
    ErrorLog "frmMDI\mnuEditFindNext_Click"
End Sub
Private Sub mnuSearchReplace_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    With frmFind
        .cmdReplace.Top = 150 'Set cmdReplace top
        .cmdReplace.Caption = "&Replace" 'Set caption
        .lblReplace.Visible = True 'Show lblReplace
        .cboReplace.Visible = True 'Show cboReplace
        .cmdReplaceAll.Visible = True 'Show cmdReplaceAll
        .Show , Me
    End With
End Sub
Private Sub mnuSearchGoTo_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    frmGoTo.Show , Me
End Sub

'// VIEW MENU
Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    SB.Visible = mnuViewStatusBar.Checked
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Status Bar", mnuViewStatusBar.Checked
End Sub
Private Sub mnuViewRuler_Click()
    mnuViewRuler.Checked = Not mnuViewRuler.Checked
    cbrRuler.Visible = mnuViewRuler.Checked
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Ruler", mnuViewRuler.Checked
End Sub
Private Sub mnuViewFileTree_Click()
    mnuViewFileTree.Checked = Not mnuViewFileTree.Checked
    picFileBar.Visible = mnuViewFileTree.Checked
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "File Tree", mnuViewFileTree.Checked
End Sub
Private Sub mnuViewFullScreen_Click()
    On Error GoTo FullScreenError
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    
    ' Select all text
    mnuEditSelectAll_Click
    ' Copy text
    SendMessage ActiveForm.rtfText.hWnd, WM_COPY, 0, 0&
    ' Paste text
    SendMessage frmFScreen.FSRTB.hWnd, WM_PASTE, 0, 0&
    ' Show full screen
    frmFScreen.Show , Me
    frmFScreen.FSRTB.SelStart = 0
FullScreenError:
    ErrorLog "frmMDI\mnuViewFullScreen_Click"
End Sub
Private Sub mnuViewToolbarEdit_Click()
    mnuViewToolbarEdit.Checked = Not mnuViewToolbarEdit.Checked
    cbrBar.Bands(5).Visible = mnuViewToolbarEdit.Checked
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Edit Toolbar", mnuViewToolbarEdit.Checked
End Sub
Private Sub mnuViewToolbarFile_Click()
    mnuViewToolbarFile.Checked = Not mnuViewToolbarFile.Checked
    cbrBar.Bands(4).Visible = mnuViewToolbarFile.Checked
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "File Toolbar", mnuViewToolbarFile.Checked
End Sub
Private Sub mnuViewToolbarFont_Click()
    mnuViewToolbarFont.Checked = Not mnuViewToolbarFont.Checked
    cbrBar.Bands(2).Visible = mnuViewToolbarFont.Checked
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Font Toolbar", mnuViewToolbarFont.Checked
End Sub
Private Sub mnuViewToolbarFormat_Click()
    mnuViewToolbarFormat.Checked = Not mnuViewToolbarFormat.Checked
    cbrBar.Bands(3).Visible = mnuViewToolbarFormat.Checked
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Format Toolbar", mnuViewToolbarFormat.Checked
End Sub
Private Sub mnuViewToolbarStandard_Click()
    mnuViewToolbarStandard.Checked = Not mnuViewToolbarStandard.Checked
    cbrBar.Bands(1).Visible = mnuViewToolbarStandard.Checked
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Standard Toolbar", mnuViewToolbarStandard.Checked
End Sub
Private Sub mnuViewToolbarWindow_Click()
    mnuViewToolbarWindow.Checked = Not mnuViewToolbarWindow.Checked
    cbrBar.Bands(6).Visible = mnuViewToolbarWindow.Checked
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Window Toolbar", mnuViewToolbarWindow.Checked
End Sub
Private Sub mnuViewStayonTop_Click()
    mnuViewStayonTop.Checked = Not mnuViewStayonTop.Checked
    If mnuViewStayonTop.Checked Then
        OnTop Me 'Put ElitePad on top
        ' Save to registry
        RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Stay On Top", 1
    Else
        NotOnTop Me 'Remove ElitePAd from top
        ' Save to registry
        RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Stay On Top", 0
    End If
End Sub
Private Sub mnuViewMode_Click(Index As Integer)
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    Dim I As Integer
    SetViewMode Index 'Set selected view mode
    For I = 0 To 2 'Uncheck all items
        mnuViewMode(I).Checked = False
    Next
    mnuViewMode(Index).Checked = True 'Check selected item
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "ViewMode", Str(Index)
End Sub
Private Sub mnuViewDocumentProperties_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    frmDocInfo.Show , Me
End Sub

'// INSERT MENU
Private Sub mnuInsertTimeDate_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    frmTimeDate.Show , Me
End Sub
Private Sub mnuInsertPicture_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    On Error GoTo PictureError
    
    CmDlg.DialogTitle = "Select Picture..."
    CmDlg.Filter = "Bitmaps (*.bmp;*.dib)|*.bmp;*.dib|GIF Images (*.gif)|*.gif|JPEG Images (*.jpg)|*.jpg|"
    CmDlg.ShowOpen
    
    'Load picture into picInsert
    picInsert.Picture = LoadPicture(CmDlg.cFileName(1))
    
    'Copy the picture into the clipboard.
    Clipboard.Clear
    Clipboard.SetData picInsert.Picture
    
    'Paste the picture into the RichTextBox.
    SendMessage ActiveForm.rtfText.hWnd, WM_PASTE, 0, 0&
PictureError:
    ErrorLog "frmMDI\mnuInsertPicture_Click"
End Sub
Private Sub mnuInsertTextFile_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    On Error GoTo InsertError
    Dim fType As String

    CmDlg.DialogTitle = "Select File to Insert" 'Set title
    CmDlg.Filter = epFilter 'Set filter
    CmDlg.CancelError = True
    CmDlg.ShowOpen 'Show open dialog
            
    ' Get file extension
    Select Case UCase(Right(CmDlg.cFileTitle(1), 3))
        Case "RTF"
            fType = rtfRTF
        Case Else
            fType = rtfText
    End Select

    rtfTemp.LoadFile CmDlg.cFileName(1), fType 'Load file into rtfTemp
    rtfTemp.SelStart = 0 'Set selStart to 0
    rtfTemp.SelLength = Len(rtfTemp.Text) 'Select all text
    SendMessage rtfTemp.hWnd, WM_CUT, 0, 0& 'Cut text from rtfTemp
    ' Paste text into rtfText
    ActiveForm.rtfText.SelText = SendMessage(ActiveForm.rtfText.hWnd, WM_PASTE, 0, 0&)
InsertError:
    If Err.Number = 32755 Then Exit Sub
    ErrorLog "frmMDI\mnuInsertTextFile_Click"
End Sub
Private Sub mnuInsertPathandFile_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    
    If Left(ActiveForm.Caption, 8) = "Document" Then 'Dont insert if doesn't exists
        Exit Sub
    Else
        ActiveForm.rtfText.SelText = ActiveForm.Caption 'Insert path anf file
    End If
End Sub
Private Sub mnuInsertSymbols_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    frmSymbols.Show , Me
End Sub

'// FORMAT MENU
Private Sub mnuFormatBullet_Click()
    Bullet
End Sub
Private Sub mnuFormatFont_Click()
    On Error GoTo FontError
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    
    With ActiveForm.rtfText
        CmDlg.CancelError = True
        CmDlg.ShowFont 'Show dialog
        .SelBold = CmDlg.FontBold 'Set bold
        .SelColor = CmDlg.FontColor 'Set font color
        .SelFontName = CmDlg.FontName 'Set fontname
        .SelFontSize = CmDlg.FontSize 'Set font size
        .SelItalic = CmDlg.FontItalic 'Set italic
        .SelStrikeThru = CmDlg.FontStrikeThru 'Set strikethru
        .SelUnderline = CmDlg.FontUnderline 'Set underline
    End With
FontError:
    If Err.Number = 32755 Then Exit Sub 'If canceled then exit sub
    ErrorLog "frmMDI/mnuFormatFont_click"
End Sub
Private Sub mnuFormatUpper_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    ActiveForm.rtfText.SelText = UCase(ActiveForm.rtfText.SelText)
End Sub
Private Sub mnuFormatLower_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    ActiveForm.rtfText.SelText = LCase(ActiveForm.rtfText.SelText)
End Sub
Private Sub mnuFormatScriptNoScript_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    ActiveForm.rtfText.SelCharOffset = 0
End Sub
Private Sub mnuFormatScriptSubscript_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    ActiveForm.rtfText.SelCharOffset = -55
End Sub
Private Sub mnuFormatScriptSuperScript_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    ActiveForm.rtfText.SelCharOffset = 55
End Sub
Private Sub mnuFormatIndent_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    'Set the forms scale mode to Millimeters
    ActiveForm.ScaleMode = vbMillimeters
    'Change the indent
    ActiveForm.rtfText.SelIndent = ActiveForm.rtfText.SelIndent + 13
    'Return form scale mode to Twips
    ActiveForm.ScaleMode = vbTwips
End Sub
Private Sub mnuFormatOutdent_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    'Set the forms scale mode to Millimeters
    ActiveForm.ScaleMode = vbMillimeters
    'Change the indent
    ActiveForm.rtfText.SelIndent = ActiveForm.rtfText.SelIndent - 13
    'Return form scale mode to Twips
    ActiveForm.ScaleMode = vbTwips
End Sub

'// TOOLS MENU
Private Sub mnuToolsOptions_Click()
    frmOptions.Show , Me
End Sub

'// WINDOW MENU
Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub
Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub
Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub
Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub
Private Sub mnuWindowNewWindow_Click()
    CreateNewDocument 'Call CreateNewDocument function
End Sub
Private Sub mnuWindowMinimizeAll_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    Dim I As Integer
    ' Minimize all documents
    For I = 1 To Forms.Count - 1
        Forms(I).WindowState = vbMinimized
    Next I
End Sub
Private Sub mnuWindowRestoreAll_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "ElitePad", 1, True: Exit Sub
    Dim I As Integer
    ' Restore all documents
    For I = 1 To Forms.Count - 1
        Forms(I).WindowState = vbNormal
    Next I
End Sub

'// HELP MENU
Private Sub mnuHelpContents_Click()
    HHShowContents Me.hWnd
End Sub
Private Sub mnuHelpIndex_Click()
    HHShowIndex Me.hWnd
End Sub
Private Sub mnuHelpSearch_Click()
    HHShowSearch Me.hWnd
End Sub
Private Sub mnuHelpTipoftheDay_Click()
    frmTip.Show , Me
End Sub
Private Sub mnuHelpAbout_Click()
    frmAbout.Show , Me
End Sub

'// POPUP MENU
Private Sub mnuPopUndo_Click()
    'mnueditundo_click
End Sub
Private Sub mnuPopCut_Click()
    mnuEditCut_Click
End Sub
Private Sub mnuPopCopy_Click()
    mnuEditCopy_Click
End Sub
Private Sub mnuPopPaste_Click()
    mnuEditPaste_Click
End Sub
Private Sub mnuPopDelete_Click()
    mnuEditDelete_Click
End Sub
Private Sub mnuPopSelectAll_Click()
    mnuEditSelectAll_Click
End Sub
Private Sub mnuPopCaseLower_Click()
    mnuFormatLower_Click
End Sub
Private Sub mnuPopCaseUpper_Click()
    mnuFormatUpper_Click
End Sub
