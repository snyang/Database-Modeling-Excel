Attribute VB_Name = "basClipboard"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2012, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

Public Function CopyToClipboard(ByRef Text As String)
    Dim objDataObject As DataObject
    '-- Copy result to clipboard
    
    Set objDataObject = New DataObject
    objDataObject.Clear
    objDataObject.SetText Text
    objDataObject.PutInClipboard
    Set objDataObject = Nothing

'        If CBool(OpenClipboard(0)) Then
'            EmptyClipboard
'            Dim hMemHandle As Long, lpData As Long, length As Long
'            length = GetStringLen(G_Output_Content)
'            hMemHandle = GlobalAlloc(0, length)
'            If CBool(hMemHandle) Then
'                lpData = GlobalLock(hMemHandle)
'                If lpData <> 0 Then
'                    CopyMemory ByVal lpData, ByVal G_Output_Content, length
'                    GlobalUnlock hMemHandle
'                    EmptyClipboard
'                    SetClipboardData 1, hMemHandle
'                End If
'                GlobalFree hMemHandle
'            End If
'            Call CloseClipboard
'        End If
End Function


Public Function GetFromClipboard() As String
    Dim objDataObject As DataObject
    Set objDataObject = New DataObject
    objDataObject.GetFromClipboard
    
    GetFromClipboard = objDataObject.GetText
    Set objDataObject = Nothing
End Function

