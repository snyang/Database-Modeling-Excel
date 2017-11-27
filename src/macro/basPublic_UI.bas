Attribute VB_Name = "basPublic_UI"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2012, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

Public Sub SelectAllListBoxItems(ByVal controlListBox As Object, _
                            ByVal selected As Boolean)
    Dim index      As Integer
    
    '-- Select all items or deselect all items
    With controlListBox
        For index = 0 To .ListCount - 1
            .selected(index) = selected
        Next
    End With
End Sub

Public Sub SelectAllText(ByVal control As Object)
    If VBA.typeName(control) = "TextBox" Then
        With control
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    ElseIf VBA.typeName(control) = "CboBox" Then
        With control
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
End Sub

Public Function MsgBoxEx_Error() As VbMsgBoxResult
    MsgBoxEx_Error = MsgBox("Error: " & Err.Description, vbInformation + vbOKOnly, App_Name)
End Function
