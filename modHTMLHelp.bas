Attribute VB_Name = "modHTMLHelp"
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

Public Const HH_HELP_CONTEXT = &HF
Private Const HH_DISPLAY_TOPIC = &H0
Private Const HH_DISPLAY_TOC = &H1
Private Const HH_DISPLAY_INDEX = &H2
Private Const HH_DISPLAY_SEARCH = &H3
Private Const HH_DISPLAY_TEXT_POPUP = &HE

Private Type tagHH_FTS_QUERY
    cbStruct          As Long
    fUniCodeStrings   As Long
    pszSearchQuery    As String
    iProximity        As Long
    fStemmedSearch    As Long
    fTitleOnly        As Long
    fExecute          As Long
    pszWindow         As String
End Type

Public Declare Function HTMLHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Private Declare Function HTMLHelpCallSearch Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByRef dwData As tagHH_FTS_QUERY) As Long

'----------SHOW HTMLHELP SEARCH----------'
Public Function HHShowSearch(lhWnd As Long)
    Dim HHFQ As tagHH_FTS_QUERY
    With HHFQ
        .cbStruct = Len(HHFQ)
        .fUniCodeStrings = 0&
        .pszSearchQuery = ""
        .iProximity = 0&
        .fStemmedSearch = 0&
        .fTitleOnly = 0&
        .fExecute = 1&
        .pszWindow = ""
    End With
    HTMLHelpCallSearch lhWnd, App.Path & "\ElitePad.chm" & ">Main", HH_DISPLAY_SEARCH, HHFQ
End Function

'----------SHOW HTMLHELP CONTENTS----------'
Public Function HHShowContents(lhWnd As Long)
    HTMLHelp lhWnd, App.Path & "\ElitePad.chm" & ">Main", HH_DISPLAY_TOC, 0
End Function

'----------SHOW HTMLHELP INDEX----------'
Public Function HHShowIndex(lhWnd As Long)
    HTMLHelp lhWnd, App.Path & "\ElitePad.chm" & ">Main", HH_DISPLAY_INDEX, 0
End Function

'----------SHOW HTMLHELP TOPIC----------'
Public Function HHShowTopic(lhWnd As Long, lngTopicID As Long)
    HTMLHelp lhWnd, App.Path & "\ElitePad.chm" & ">Main", HH_HELP_CONTEXT, lngTopicID
End Function
