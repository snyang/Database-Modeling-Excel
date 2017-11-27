Attribute VB_Name = "basWorkbookEventHandler"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2012, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

Public Sub basWorkbookEventHandler_Activate()
    On Error Resume Next
    Call AddCommandBar
End Sub

Public Sub basWorkbookEventHandler_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Set barDBModeling = CommandBars.Item(BAR_NAME)
    If Not barDBModeling Is Nothing Then
        Call SetCommandBarButtonsToNothing
        barDBModeling.Delete
        Set barDBModeling = Nothing
    End If
End Sub

Public Sub basWorkbookEventHandler_WindowActivate(ByVal Wn As Excel.Window)
    On Error Resume Next
    Call AddCommandBar
End Sub

Public Sub basWorkbookEventHandler_WindowDeactivate(ByVal Wn As Excel.Window)
    On Error Resume Next
    Call DeleteCommandBar
End Sub



