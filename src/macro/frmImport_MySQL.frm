VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImport_MySQL 
   Caption         =   "Import"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6075
   OleObjectBlob   =   "frmImport_MySQL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImport_MySQL"
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

'-----------------------------------------------------------
'           MySQL
'-----------------------------------------------------------

Private mDatabaseType As String
Public Property Get DatabaseType() As String
        DatabaseType = mDatabaseType
End Property
Public Property Let DatabaseType(Value As String)
        mDatabaseType = Value
End Property

Private Sub btnRefreshDatabase_Click()
    Dim conn As ADODB.Connection
    Dim sSQL As String
    Dim oRs As ADODB.Recordset
    Dim lastIndex As Integer
    Dim index As Integer

    If Len(Trim(txtServer.Text)) = 0 Then
        MsgBox "Please input " & labServer.Caption & ".", vbOKOnly + vbInformation, App_Name
        txtServer.SetFocus
        Exit Sub
    End If

    On Error GoTo Flag_Err

    Set conn = basImport_MySQL.CreateConnection(Trim(txtServer.Text), _
                            "", _
                            Trim(txtUser.Text), _
                            Trim(txtPassword.Text), _
                            Trim(cboDriver.Text), _
                            Trim(txtPort.Text))
    conn.Open
    sSQL = "SELECT * FROM information_schema.SCHEMATA S ORDER BY S.SCHEMA_NAME;"

    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    lastIndex = -1
    index = -1
    cboDatabase.Clear
    Do While Not oRs.EOF
        index = index + 1
        cboDatabase.AddItem (oRs("SCHEMA_NAME"))
        If oRs("SCHEMA_NAME") = basImport_MySQL.SERVER_DATABASE_NAME Then
            lastIndex = index
        End If
        '-- Move next record
        oRs.MoveNext
    Loop

    If cboDatabase.ListCount > 0 Then
        If lastIndex > 0 Then
            cboDatabase.ListIndex = lastIndex
        Else
            cboDatabase.ListIndex = 0
        End If
    End If
    '-- Close record set
    oRs.Close
    conn.Close
    Set oRs = Nothing
    Set conn = Nothing

    '-- Set public variant
    basImport_MySQL.SERVER_NAME = txtServer.Text
    basImport_MySQL.DRIVER_NAME = cboDriver.Text
    basImport_MySQL.PORT_ID = txtPort.Text
    Exit Sub
Flag_Err:
    If Not oRs Is Nothing Then oRs.Close
    Set oRs = Nothing
    Set conn = Nothing

    Call MsgBoxEx_Error
End Sub

Private Sub btnRefreshTable_Click()
    Dim conn As ADODB.Connection
    Dim sSQL As String
    Dim oRs As ADODB.Recordset
    Dim lastIndex As Integer
    Dim index As Integer
    Dim sTableName As String

    On Error GoTo Flag_Err
    If Cells.Item(Table_Sheet_Row_TableName, _
                    Table_Sheet_Col_TableName).Text = "" Then
        sTableName = basImport_MySQL.SERVER_TABLE_NAME
    Else
        sTableName = Cells.Item(Table_Sheet_Row_TableName, Table_Sheet_Col_TableName).Text
    End If

    Set conn = basImport_MySQL.CreateConnection(Trim(txtServer.Text), _
                            Trim(cboDatabase.Text), _
                            Trim(txtUser.Text), _
                            Trim(txtPassword.Text), _
                            Trim(cboDriver.Text), _
                            Trim(txtPort.Text))
    conn.Open
    sSQL = "SELECT * FROM information_schema.`TABLES` T" _
        & Line & "WHERE T.TABLE_SCHEMA = DATABASE()" _
        & Line & "ORDER BY T.TABLE_NAME;"

    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    lastIndex = -1
    index = -1
    lstTables.Clear
    Do While Not oRs.EOF
        index = index + 1
        lstTables.AddItem (oRs("TABLE_NAME"))
        If oRs("TABLE_NAME") = sTableName Then
            lastIndex = index
        End If
        '-- Move next record
        oRs.MoveNext
    Loop

    If lstTables.ListCount > 0 Then
        If lastIndex > 0 Then
            lstTables.ListIndex = lastIndex
        Else
            lstTables.ListIndex = 0
        End If
    End If

    '-- Set public variant
    basImport_MySQL.SERVER_DATABASE_NAME = cboDatabase.Text
    '-- Close record set
    oRs.Close
    conn.Close
    Set oRs = Nothing
    Set conn = Nothing

    Exit Sub
Flag_Err:
    If Not oRs Is Nothing Then oRs.Close
    If Not conn Is Nothing Then conn.Close
    Set oRs = Nothing
    Set conn = Nothing

    Call MsgBoxEx_Error
End Sub

Private Sub Init()
    Dim iActiveSheet As Integer
    Dim index As Integer
    Dim iSheet As Integer

    Dim shtCurrent As Worksheet
    On Error GoTo F_Error
    
    DatabaseType = DBName_MySQL
    
    '-- Init
    Me.Caption = "Import from " & DatabaseType
    
    '-- Init Driver
    Call cboDriver.AddItem("{MySQL ODBC 5.2 UNICODE Driver}")
    Call cboDriver.AddItem("{MySQL ODBC 5.1 Driver}")
    Call cboDriver.AddItem("{MySQL ODBC 3.51 Driver}")
    cboDriver.ListIndex = 0
    If Len(basImport_MySQL.DRIVER_NAME) > 0 Then
        cboDriver.Text = basImport_MySQL.DRIVER_NAME
    End If
    
    '-- Init Server
    If Len(basImport_MySQL.PORT_ID) > 0 Then
        txtPort.Text = basImport_MySQL.PORT_ID
    Else
        txtPort.Text = "3306"
    End If
    If Len(basImport_MySQL.SERVER_NAME) > 0 Then
        txtServer.Text = basImport_MySQL.SERVER_NAME
    Else
        txtServer.Text = "localhost"
    End If
    txtServer.SetFocus

    '-- init combo sheet
    cboSheet.Clear
    index = 0
    iActiveSheet = -1
    iActiveSheet = ThisWorkbook.ActiveSheet.index - Sheet_First_Table + 1
    For iSheet = Sheet_First_Table To ThisWorkbook.Sheets.Count
        Set shtCurrent = ThisWorkbook.Sheets(iSheet)
        '-- Set Caption = index & tablecaption
        cboSheet.AddItem shtCurrent.Name
        cboSheet.List(index, 1) = shtCurrent.index

        If ThisWorkbook.ActiveSheet.index = shtCurrent.index Then
            iActiveSheet = index
        End If
        index = index + 1
    Next

    If cboSheet.ListCount > 0 Then
        If iActiveSheet >= 0 Then
            cboSheet.ListIndex = iActiveSheet
        Else
            cboSheet.ListIndex = 0
        End If
    End If
    
    '-- set focus
    txtServer.SetFocus

F_Error:
    Set shtCurrent = Nothing
End Sub

Private Sub DoImport()
    On Error GoTo Flag_Err

    Dim index               As Integer
    Dim shtTemplate         As Worksheet
    Dim shtCurrent          As Worksheet
    Dim currentSheetIndex   As Integer
    Dim conn                As ADODB.Connection
    Dim clearSheet          As Boolean
    Dim TableName           As String
    Dim table               As clsLogicalTable
    Dim isSetPublicVarient  As Boolean
    isSetPublicVarient = False

    If cboSheet.ListIndex < 0 Then Exit Sub

    clearSheet = (Me.chkClearSheet.Value)
    Set conn = basImport_MySQL.CreateConnection(Trim(txtServer.Text), _
                            cboDatabase.Text, _
                            Trim(txtUser.Text), _
                            Trim(txtPassword.Text), _
                            Trim(cboDriver.Text), _
                            Trim(txtPort.Text))
    conn.Open
    
    currentSheetIndex = CInt(cboSheet.List(cboSheet.ListIndex, 1))
    Set shtTemplate = ThisWorkbook.Sheets(currentSheetIndex)
    
    For index = 0 To Me.lstTables.ListCount - 1
        If lstTables.selected(index) Then
            '-- Get Table Definition
            TableName = lstTables.List(index, 0)
            
            '-- Get a sheet which is used to store table information
            
            If Me.chkAutoMatch.Value Then
                Set shtCurrent = GetSheetFromTableName(TableName)
                clearSheet = False
                If shtCurrent Is Nothing Then
                    Set shtCurrent = CopyASheet(shtTemplate, , ThisWorkbook.Sheets(currentSheetIndex))
                    currentSheetIndex = shtCurrent.index
                    clearSheet = True
                End If
            Else
                If Me.chkCreateASheet.Value Then
                    currentSheetIndex = ThisWorkbook.Sheets.Count
                    Set shtCurrent = CopyASheet(shtTemplate, , ThisWorkbook.Sheets(currentSheetIndex))
                Else
                    Set shtCurrent = ThisWorkbook.Sheets(currentSheetIndex)
                End If
            End If
            
            '-- Set public variant
            If isSetPublicVarient = False Then
                basImport_MySQL.SERVER_TABLE_NAME = TableName
                isSetPublicVarient = True
            End If
            Set table = basImport_MySQL.GetLogicalTable(conn, TableName)
            '-- Write to sheet
            shtCurrent.Select
            Call basTableSheet.SetTableInfoToWorksheet(shtCurrent, table, clearSheet)
        End If
    Next

    '-- Close connection
    conn.Close
    Set conn = Nothing
    
    Exit Sub
Flag_Err:
    Set conn = Nothing

    Call MsgBoxEx_Error
End Sub

Private Sub btnSelectAllTable_Click()
    Call SelectAllListBoxItems(lstTables, True)
End Sub

Private Sub btnUnSelectAllTable_Click()
    Call SelectAllListBoxItems(Me.lstTables, False)
End Sub

Private Sub cboDatabase_Enter()
    Call SelectAllText(cboDatabase)
End Sub

Private Sub cboDriver_Enter()
    Call SelectAllText(cboDriver)
End Sub

Private Sub cboSheet_Enter()
    Call SelectAllText(cboSheet)
End Sub

Private Sub txtPort_Enter()
    Call SelectAllText(txtPort)
End Sub

Private Sub txtServer_Enter()
    Call SelectAllText(txtServer)
End Sub

Private Sub chkAutoMatch_Change()
    If Me.chkAutoMatch.Value Then
        Me.chkCreateASheet.Value = True
        Me.chkClearSheet.Value = True
    End If
End Sub

Private Sub txtPassword_Enter()
    Call SelectAllText(txtPassword)
End Sub

Private Sub txtUser_Enter()
    Call SelectAllText(txtUser)
End Sub

Private Sub btnImport_Click()
    Call DoImport
End Sub

Private Sub UserForm_Initialize()
    Call Init
End Sub

Private Sub chkCreateASheet_Change()
    If Me.chkCreateASheet.Value Then
        Me.chkClearSheet.Value = True
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
