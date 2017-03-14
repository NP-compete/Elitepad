Attribute VB_Name = "modMenuStatus"
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

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal LParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long

Private Const MF_BYPOSITION = &H400&
Private Const MF_HILITE = &H80&
Private Const WM_MENUSELECT = &H11F
Private Const GWL_WNDPROC = -4

Public lpPrevWndProc As Long
Public gHW As Long

Private Function TMenu(MenuCaption As String) As String

    ' Description for File menu
    If Left(MenuCaption, 4) = "&New" Then TMenu = "Create a new document"
    If Left(MenuCaption, 5) = "&Open" Then TMenu = "Open an existing document"
    If Left(MenuCaption, 6) = "&Close" Then TMenu = "Close the active document"
    If Left(MenuCaption, 10) = "Clos&e All" Then TMenu = "Close all documents"
    If Left(MenuCaption, 5) = "&Save" Then TMenu = "Save the active document"
    If Left(MenuCaption, 8) = "Save &As" Then TMenu = "Save the active document with a new name"
    If Left(MenuCaption, 9) = "Save A&ll" Then TMenu = "Save all documents"
    If Left(MenuCaption, 18) = "Sa&ve Selection As" Then TMenu = "Save selection as file..."
    If Left(MenuCaption, 16) = "Revert to Save&d" Then TMenu = "Revert the current document to the last saved version"
    If Left(MenuCaption, 13) = "Mana&ge Files" Then TMenu = "Copy, rename or delete files"
    If Left(MenuCaption, 6) = "&Print" Then TMenu = "Print the active document"
    If Left(MenuCaption, 5) = "E&xit" Then TMenu = "Quit the application; prompts to save documents"
    
    ' Description for Edit menu
    If Left(MenuCaption, 5) = "&Undo" Then TMenu = "Undo your last action"
    If Left(MenuCaption, 5) = "&Redo" Then TMenu = "Redo the previously undone action"
    If Left(MenuCaption, 4) = "Cu&t" Then TMenu = "Cut the selection to the Clipboard"
    If Left(MenuCaption, 5) = "&Copy" Then TMenu = "Copy the selection to the Clipboard"
    If Left(MenuCaption, 6) = "&Paste" Then TMenu = "Insert Clipboard contents"
    If Left(MenuCaption, 7) = "&Delete" Then TMenu = "Delete the selected text"
    If Left(MenuCaption, 11) = "&Select All" Then TMenu = "Select the entire document"
    If Left(MenuCaption, 11) = "Mar&k Clean" Then TMenu = "Mark the active document as unmodified"

    ' Description for Search menu
    If Left(MenuCaption, 5) = "&Find" Then TMenu = "Search for text in the active document"
    If Left(MenuCaption, 10) = "Find &Next" Then TMenu = "Find next occurrence of search string"
    If Left(MenuCaption, 8) = "&Replace" Then TMenu = "Replace occurrences of search string"
    If Left(MenuCaption, 6) = "&Go To" Then TMenu = "Go to specified line"
    
    ' Description for View menu
    If Left(MenuCaption, 11) = "&Status Bar" Then TMenu = "Show or hide the status bar"
    If Left(MenuCaption, 6) = "&Ruler" Then TMenu = "Show or hide the ruler"
    If Left(MenuCaption, 15) = "Fi&le Tree View" Then TMenu = "Show or hide the file tree"
    If Left(MenuCaption, 12) = "&Full Screen" Then TMenu = "Toggle full screen view"
    If Left(MenuCaption, 8) = "&No Wrap" Then TMenu = "Set view mode"
    If Left(MenuCaption, 10) = "&Word Wrap" Then TMenu = "Set view mode"
    If Left(MenuCaption, 8) = "WY&SIWYG" Then TMenu = "Set view mode"
    If Left(MenuCaption, 12) = "Stay on &Top" Then TMenu = "Make ElitePad stay on top"
    If Left(MenuCaption, 20) = "Document P&roperties" Then TMenu = "View properties for this document"
    
    ' Description for Insert menu
    If Left(MenuCaption, 14) = "Time and &Date" Then TMenu = "Insert current time and date"
    If Left(MenuCaption, 8) = "&Picture" Then TMenu = "Insert picture from file"
    If Left(MenuCaption, 10) = "&Text File" Then TMenu = "Insert file into active document"
    If Left(MenuCaption, 14) = "Path and &File" Then TMenu = "Insert path and filename"
    If Left(MenuCaption, 8) = "&Symbols" Then TMenu = "Insert ANSI symbols"
    
    ' Description for Format menu
    If Left(MenuCaption, 5) = "&Font" Then TMenu = "Set selected text font"
    If Left(MenuCaption, 7) = "&Bullet" Then TMenu = "Set bullets"
    If Left(MenuCaption, 14) = "To &Upper Case" Then TMenu = "Change the selection to upper case"
    If Left(MenuCaption, 14) = "To &Lower Case" Then TMenu = "Change the selection to lower case"
    If Left(MenuCaption, 16) = "Increase &Indent" Then TMenu = "Increase indentation of selected lines"
    If Left(MenuCaption, 14) = "&Reduce Indent" Then TMenu = "Reduce indentation of selected lines"

    ' Description for Tools menu
    If Left(MenuCaption, 8) = "&Options" Then TMenu = "Show the Options dialog"

    ' Description for Window menu
    If Left(MenuCaption, 8) = "&Cascade" Then TMenu = "Arrange windows so they overlap"
    If Left(MenuCaption, 16) = "Tile &Horizontal" Then TMenu = "Arrange windows as tiles down the screen"
    If Left(MenuCaption, 14) = "Tile &Vertical" Then TMenu = "Arrange windows as tiles across the screen"
    If Left(MenuCaption, 14) = "&Arrange Icons" Then TMenu = "Arrange icons at the bottom of the window"
    If Left(MenuCaption, 13) = "&Minimize All" Then TMenu = "Minimize windows"
    If Left(MenuCaption, 12) = "&Restore All" Then TMenu = "Restore windows to normal position"
    
    ' Description for Help menu
    If Left(MenuCaption, 15) = "Tip of the &Day" Then TMenu = "Display a tip of the day"
    If Left(MenuCaption, 6) = "&About" Then TMenu = "Display program information, version number and copyright"
    
End Function

Public Sub Hook()
    ' Begin hooking into messages.
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook()
    ' Cease hooking into messages.
    SetWindowLong gHW, GWL_WNDPROC, lpPrevWndProc
End Sub

Function AnyLit(hSubSubMenu As Long) As Long
    Dim I As Long
    Dim MenuCount As Long

    ' Get the number of items in the menu.
    MenuCount = GetMenuItemCount(hSubSubMenu)

    ' Loop through the menu items.
    For I = 0 To MenuCount - 1
        ' Check whether this item is highlighted.
        If GetMenuState(hSubSubMenu, I, MF_BYPOSITION) And MF_HILITE Then
            AnyLit = True
            Exit Function
        End If
    Next I

    ' Return FALSE, no items highlighted.
    AnyLit = False
End Function

Private Sub WalkSubMenu(hSubMenu As Long)
    Dim I As Long
    Dim MenuItems As Long
    Dim hSubSubMenu As Long
    Dim Buffer As String
    Dim Result As Long

    ' Get the count of menu items in this menu.
    MenuItems = GetMenuItemCount(hSubMenu)
    ' Loop through all the items on the menu.
    For I = 0 To MenuItems - 1
        ' Determine whether this item is highlighted.
        If GetMenuState(hSubMenu, I, MF_BYPOSITION) And MF_HILITE Then
            ' Attempt to get a submenu for each menu item.
            hSubSubMenu = GetSubMenu(hSubMenu, I)

            ' Check for a submenu with something selected on it.
            If hSubSubMenu And AnyLit(hSubSubMenu) Then
                ' There is a submenu with a selection so walk it.
                WalkSubMenu hSubSubMenu
            Else    'This is it.
                ' Set buffer size.
                Buffer = Space(255)

                ' Call the API to get the caption for the menu item.
                Result = GetMenuString(hSubMenu, I, Buffer, Len(Buffer), MF_BYPOSITION)
                ' Trim the buffer of extra characters.
                Buffer = Left$(Buffer, Result)
                frmMDI.SB.Panels(1).Text = TMenu(Buffer)
                ' Exit this Sub procedure.
                Exit Sub
            End If
        End If
    Next I
End Sub

Public Sub FindHilite(TheForm As Form)
    Dim hMenu As Long
    Dim hSubMenu As Long
    Dim I As Long
    Dim MenuCount As Long

    ' Clear any previous description.
    frmMDI.SB.Panels(1).Text = "Press F1 for help."

    ' Get the menu handle.
    hMenu = GetMenu(TheForm.hWnd)

    ' Check to see if there is no menu.
    If hMenu <> 0 Then
        ' Get the number of top-level menus.
        MenuCount = GetMenuItemCount(hMenu)

        ' Enumerate through all top-level menus.
        For I = 0 To MenuCount - 1
            ' Ignore top-level menus not currently selected.
            If GetMenuState(hMenu, I, MF_BYPOSITION) And MF_HILITE Then
                ' Get a handle to the submenu.
                hSubMenu = GetSubMenu(hMenu, I)
                ' Walk the submenu.
                WalkSubMenu hSubMenu
            End If
        Next I
    End If
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal LParam As Long) As Long
    ' Check for a menu selection message.
    If uMsg = WM_MENUSELECT Then
        FindHilite frmMDI
    End If
    ' Pass the message to Windows for processing.
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, LParam)
End Function
