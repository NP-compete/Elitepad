Attribute VB_Name = "modMessageError"
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

Public Sub MsgE(sMessage As String, sCaption As String, Icon As Integer, bOk As Boolean)
    '** Description:
    '** Show message with custom MsgBox
    If bOk = True Then
        frmMsgBox.cmdOk.Visible = True
    Else
        frmMsgBox.cmdCancel.Visible = True
        frmMsgBox.cmdNo.Visible = True
        frmMsgBox.cmdYes.Visible = True
    End If
    ' See wich icon is
    If Icon = 0 Then 'Critical
        frmMsgBox.imgCri.Visible = True
    Else 'Help
        frmMsgBox.imgHelp.Visible = True
    End If
    frmMsgBox.Caption = sCaption 'Set msgbox caption
    frmMsgBox.lblMsg.Caption = sMessage 'Set message
    frmMsgBox.Show vbModal, frmMDI 'Show form
End Sub

Public Function ErrorLog(Proc As String)
    '** Description:
    '** Error loger
    Dim ErrDes As String
    Dim ErrNum As Long
    
    If Err.Number = 0 Then Exit Function
    ErrDes = Err.Description 'Set error description
    ErrNum = Err.Number 'Set error number
    
    ' Open Errorlog.log and log the error
    Open App.Path & "\ErrorLog.log" For Append As #1
        Print #1, _
         " " & vbCrLf & _
        "Description = " & ErrDes & vbCrLf & _
        "     Number = " & ErrNum & vbCrLf & _
        "     Source = " & Proc & vbCrLf & _
        "       Time = " & Now & vbCrLf & _
        " " & vbCrLf & _
        "----------------------------"
    Close #1
    ' Show message with the error
    MsgE "There was a error in : " & Proc & vbCrLf & _
         "Error # : " & ErrNum & vbCrLf & _
         "Description : " & ErrDes, "ElitePad - ERROR", 0, True
End Function

