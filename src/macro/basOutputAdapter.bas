Attribute VB_Name = "basOutputAdapter"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2012, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

Enum enmOutputMode
    opmToClipboard = 0
    opmToExcelSheet
    opmToFile
End Enum

Const G_Output_SHEET_NAME = "Output"

Private mOutputOptions              As clsOutputOptions
Private outputContentArray()        As String

Public Function CreateOutputOptions(Optional OutputMode As enmOutputMode = opmToClipboard, _
            Optional OutputFullName As String = "") As clsOutputOptions
    Dim outputOptions As clsOutputOptions
    Set outputOptions = New clsOutputOptions
      
    outputOptions.OutputMode = OutputMode
    outputOptions.OutputFullName = OutputFullName
    
    Set CreateOutputOptions = outputOptions
End Function

Public Function Output_Initialize(outputOptions As clsOutputOptions)
    If outputOptions Is Nothing Then
        Set mOutputOptions = CreateOutputOptions
    Else
        Set mOutputOptions = outputOptions
    End If
    
    ReDim outputContentArray(0) As String
    ReDim objOutputLine(1) As Long
End Function

Public Function Output_Write(ByVal Text As String, Optional outputID As Integer = 0)
    Dim outputContent As String
    
    outputContent = GetOuputContent(outputID)
    
    outputContent = outputContent & Text
    
    SetOutputContentToCollection outputID, outputContent
End Function

Public Function Output_WriteLine(ByVal Text As String, Optional outputID As Integer = 0)
    Output_Write Text & vbCrLf, outputID
End Function

Private Function GetOuputContent(outputID As Integer) As String
    If UBound(outputContentArray) < outputID Then
        ReDim Preserve outputContentArray(outputID) As String
    End If

    GetOuputContent = outputContentArray(outputID)
End Function

Private Function SetOutputContentToCollection(outputID As Integer, outputContent As String)
    outputContentArray(outputID) = outputContent
End Function

Private Function GetAllOuputContentString() As String
    Dim outputContent As String
    Dim index As Integer
    
    For index = 0 To UBound(outputContentArray)
        If index > 0 Then outputContent = outputContent & vbCrLf
        outputContent = outputContent & outputContentArray(index)
    Next
    
    GetAllOuputContentString = outputContent
End Function

Public Function Output_Copy() As String
    Dim Content As String
    
    Content = GetAllOuputContentString
    ReDim outputContentArray(0) As String
    
    Output_Copy = Content
    
    If mOutputOptions.OutputMode = opmToFile Then
        SaveToTextFile mOutputOptions.OutputFullName, Content
        Exit Function
    End If
    
    CopyToClipboard Output_Copy

    If mOutputOptions.OutputMode = opmToExcelSheet Then
        Dim sheet As Worksheet
        Set sheet = Sheets(G_Output_SHEET_NAME)
        sheet.Cells.ClearContents
        sheet.Range("A1").Select
        sheet.Paste
    End If
End Function

Public Function GetStringLen(ByRef Text As String) As Long
    Dim i As Long
    Dim length As Long

    For i = 1 To Len(Text)
        If Asc(Mid(Text, i, 1)) < 255 And Asc(Mid(Text, i, 1)) >= 0 Then
            length = length + 1
        Else
            length = length + LenB(Mid(Text, i, 1))
        End If
    Next
    
    GetStringLen = length
End Function

