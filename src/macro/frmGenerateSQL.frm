VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGenerateSQL 
   Caption         =   "Generate for <database>"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   OleObjectBlob   =   "frmGenerateSQL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGenerateSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2012, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

'-- Indicates whether the form is initialized.
Private mInitialized  As Boolean

Private mDatabaseType As String
Public Property Get DatabaseType() As String
        DatabaseType = mDatabaseType
End Property
Public Property Let DatabaseType(Value As String)
        mDatabaseType = Value
End Property

Private Sub SelectAllItem(booSel As Boolean)
    Dim index      As Integer
    
    '-- Select all items or deselect all items
    With lstTables
        For index = 0 To .ListCount - 1
            .selected(index) = booSel
        Next
    End With
End Sub

'---------------------------------------------
'-  Get all seleted table information
'---------------------------------------------
Private Function GetSelectedLogicalTables() As Collection
    Dim objLogicalTables    As Collection
    Dim index               As Integer
    Dim iSheet              As Integer
    
    Set objLogicalTables = New Collection
    
    With lstTables
        For index = 0 To .ListCount - 1
            If .selected(index) Then
                iSheet = .List(index, 1)
                Call objLogicalTables.Add(GetTableInfoFromWorksheet(ThisWorkbook.Sheets(iSheet)))
            End If
        Next
    End With
    
    '-- Return
    Set GetSelectedLogicalTables = objLogicalTables
End Function

Private Sub CopyCreateTableSQL(ByVal withComment As Boolean)
    On Error GoTo Flag_Err

    Call basPublicDatabase.GetDatabaseProvider(DatabaseType).GetSQLCreateTable(GetSelectedLogicalTables, withComment)
    
    Exit Sub
Flag_Err:
    Call MsgBoxEx_Error
End Sub

Private Sub CopyDropTableSQL()
    On Error GoTo Flag_Err
    
    Call basPublicDatabase.GetDatabaseProvider(DatabaseType).GetSQLDropTable(GetSelectedLogicalTables)
   
    Exit Sub
Flag_Err:
    Call MsgBoxEx_Error
End Sub

Private Sub CopyDropAndCreateTableSQL(ByVal withComment As Boolean)
    On Error GoTo Flag_Err
    
    Call basPublicDatabase.GetDatabaseProvider(DatabaseType).GetSQLDropAndCreateTable(GetSelectedLogicalTables, withComment)
    
    Exit Sub
Flag_Err:
    Call MsgBoxEx_Error
End Sub

Private Sub CopyCreateTableIfNotExistsSQL()
    On Error GoTo Flag_Err
    
    Call basPublicDatabase.GetDatabaseProvider(DatabaseType).GetSQLCreateTableIfNotExists(GetSelectedLogicalTables)
    
    Exit Sub
Flag_Err:
    Call MsgBoxEx_Error
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Me.MultiPage1.SelectedItem.index = 0 Then
        If Me.optCreateTableSQL.Value Then
            Call CopyCreateTableSQL(Me.chkWithComment.Value)
            
        ElseIf optDropTableSQL.Value = True Then
            Call CopyDropTableSQL
            
        ElseIf optDropAndCreateSQL.Value = True Then
            Call CopyDropAndCreateTableSQL(Me.chkWithComment.Value)
        
        ElseIf Me.optCreateTableIfNotExistsSQL.Value Then
            Call CopyCreateTableIfNotExistsSQL
            
        End If
    End If
    
    '-- Return
    Unload Me
End Sub

Private Sub cmdSelectAll_Click()
    Call SelectAllItem(True)
End Sub

Private Sub cmdSelectNone_Click()
    Call SelectAllItem(False)
End Sub

Private Sub UserForm_Activate()
    If mInitialized Then Exit Sub
    
    mInitialized = True
    Me.Caption = "Generate for " & DatabaseType
    Me.MultiPage1.Pages(0).Caption = DatabaseType
    
    Dim iSheet      As Integer
    Dim oSheet      As Worksheet
    Dim index       As Integer
    With lstTables
        '-- Get create tables's SQL
        index = 0
        For iSheet = Sheet_First_Table To ThisWorkbook.Sheets.Count
            Set oSheet = ThisWorkbook.Sheets(iSheet)
            If VBA.StrComp( _
                    LCase(TrimEx( _
                        oSheet.Cells.Item(Table_Sheet_Row_TableStatus, Table_Sheet_Col_TableStatus).Text)) _
                    , Table_Sheet_TableStatus_Ignore) _
                <> 0 Then
                .AddItem (ThisWorkbook.Sheets(iSheet).Name)
                .List(index, 1) = iSheet
                index = index + 1
            End If
        Next
    End With
    
    '-- Defaut Select ALL
    Call SelectAllItem(True)

End Sub

Private Sub UserForm_Initialize()
    mInitialized = False
End Sub
