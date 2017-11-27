Attribute VB_Name = "basUpgrade"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2014, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

Public Function ConfigureTheExcel(ByVal databaseName As String)
    Call SetTheExcelTypeVariable(databaseName)
    Call AdjustExcelSheetFocus
End Function

Private Sub SetTheExcelTypeVariable(ByVal databaseName As String)
    Dim vbCom   As VBComponent
    Dim bFound  As Boolean

    Dim i As Long
    i = 1
    bFound = False
    Do While i <= ThisWorkbook.VBProject.VBComponents.Count
        Set vbCom = ThisWorkbook.VBProject.VBComponents(i)

        If vbCom.Name = "basAppSetting" Then
            bFound = True
            Exit Do
        End If
        i = i + 1
    Loop
    
    If bFound = False Then Exit Sub
    
    Dim sLineInfo As String
    Const Statement_Need_Replace = "Public Const The_Excel_Type                     As String"
    For i = 1 To vbCom.CodeModule.CountOfLines
        If Left(vbCom.CodeModule.Lines(i, 1), Len(Statement_Need_Replace)) _
            = Statement_Need_Replace Then
            
            Select Case databaseName
            Case DBName_SQLServer
                Call vbCom.CodeModule.ReplaceLine(i, "Public Const The_Excel_Type                     As String = DBName_SQLServer")
            Case DBName_DB2
                Call vbCom.CodeModule.ReplaceLine(i, "Public Const The_Excel_Type                     As String = DBName_DB2")
            Case DBName_MariaDB
                Call vbCom.CodeModule.ReplaceLine(i, "Public Const The_Excel_Type                     As String = DBName_MariaDB")
            Case DBName_MySQL
                Call vbCom.CodeModule.ReplaceLine(i, "Public Const The_Excel_Type                     As String = DBName_MySQL")
            Case DBName_Oracle
                Call vbCom.CodeModule.ReplaceLine(i, "Public Const The_Excel_Type                     As String = DBName_Oracle")
            Case DBName_PostgreSQL
                Call vbCom.CodeModule.ReplaceLine(i, "Public Const The_Excel_Type                     As String = DBName_PostgreSQL")
            Case DBName_SQLite
                Call vbCom.CodeModule.ReplaceLine(i, "Public Const The_Excel_Type                     As String = DBName_SQLite")
            Case Else
                Call vbCom.CodeModule.ReplaceLine(i, "Public Const The_Excel_Type                     As String = DBName_All")
            End Select
            Exit For
        End If
    Next
End Sub

Private Sub AdjustExcelSheetFocus()
    Dim oSheet As Excel.Worksheet
    
    ThisWorkbook.Activate
    For Each oSheet In ThisWorkbook.Sheets
        oSheet.Activate
        oSheet.Cells(1, 1).Activate
    Next
    ThisWorkbook.Sheets(1).Activate
End Sub


