VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImport 
   Caption         =   "Import"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   OleObjectBlob   =   "frmImport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImport"
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

'-----------------------------------------------------------
'   Import Form
'-----------------------------------------------------------
Option Explicit

Const PageConnectIndex      As Integer = 0
Const PageTablesIndex       As Integer = 1
Const PageOptionIndex       As Integer = 2
Const ConnectText           As String = "Connect >"
Const NextText              As String = "Next >"
Const ImportText            As String = "Import"

Private mInitialized As Boolean
Private SkipShowTablesStatisticsInfo As Boolean
Private mImportProvider As IImportProvider

Private mDatabaseType As String
Public Property Get DatabaseType() As String
        DatabaseType = mDatabaseType
End Property
Public Property Let DatabaseType(Value As String)
        mDatabaseType = Value
End Property

Public Property Get ImportProvider() As IImportProvider
    If mImportProvider Is Nothing Then
        Set mImportProvider = basPublicDatabase.GetImportProvider(DatabaseType)
    End If
    
    Set ImportProvider = mImportProvider
End Property

Private Sub SetWizardStatus(newPageValue As Integer)
    Dim index As Integer
    
    For index = 0 To Me.MultiPageMain.Pages.Count - 1
        If index = newPageValue Then
            Me.MultiPageMain.Pages(index).Enabled = True
        Else
            Me.MultiPageMain.Pages(index).Enabled = False
        End If
    Next
    Me.MultiPageMain.Value = newPageValue
    
    If Me.MultiPageMain.Value = 0 Then
        btnPrevious.Enabled = False
    Else
        btnPrevious.Enabled = True
    End If

    If Me.MultiPageMain.Value = PageConnectIndex Then
        btnNext.Caption = ConnectText
    ElseIf Me.MultiPageMain.Value = PageTablesIndex Then
        btnNext.Caption = NextText
    ElseIf Me.MultiPageMain.Value = PageOptionIndex Then
        btnNext.Caption = ImportText
    Else
        btnNext.Caption = ConnectText
    End If
End Sub

Private Sub DoConnect()
    Dim conn As ADODB.Connection
    Dim sSQL As String
    Dim oRs As ADODB.Recordset
    Dim lastIndex As Integer
    Dim index As Integer
    Dim sTableName As String

    On Error GoTo Flag_Err
    
    If Cells.Item(Table_Sheet_Row_TableName, _
                    Table_Sheet_Col_TableName).Text = "" Then
        sTableName = Me.ImportProvider.GetOptions().LastAccessTableName
    Else
        sTableName = Cells.Item(Table_Sheet_Row_TableName, Table_Sheet_Col_TableName).Text
    End If

    Set conn = New Connection
    If Me.MultiPageConnection.Value = 0 Then
        conn.connectionString = GetConnectionString()
    Else
        conn.connectionString = txtConnectionString.Text
    End If
    
    conn.Open
    
    sSQL = Me.ImportProvider.GetTablesSql()

    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    lastIndex = -1
    index = -1
    lstTables.Clear
    Do While Not oRs.EOF
        index = index + 1
        lstTables.AddItem (oRs("name"))
        If oRs("name") = sTableName Then
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

    Call ShowTablesInfo
    '-- Close record set
    oRs.Close
    conn.Close
    Set oRs = Nothing
    Set conn = Nothing

    Call SaveConnectionOptions
    Exit Sub
Flag_Err:
    Set oRs = Nothing
    Set conn = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Function GetConnectionString() As String
    GetConnectionString = Me.ImportProvider.CreateConnectionString(Trim(cboProvider.Text), _
                            Trim(txtServer.Text), _
                            Trim(txtUser.Text), _
                            Trim(txtPassword.Text), _
                            Trim(cboDatabase.Text))
End Function

Private Sub DoImport()
    On Error GoTo Flag_Err

    Dim index               As Integer
    Dim shtTemplate         As Worksheet
    Dim shtCurrent          As Worksheet
    Dim insertSheetIndex    As Integer
    Dim conn                As ADODB.Connection
    Dim TableName           As String
    Dim table               As clsLogicalTable
    Dim isSetPublicVarient  As Boolean
    isSetPublicVarient = False

    If cboSheet.ListIndex < 0 Then Exit Sub

    Dim clearExistedData    As Boolean
    clearExistedData = Me.chkClearExistedData.Value
    
    Set conn = New Connection
    If Me.MultiPageConnection.Value = 0 Then
        conn.connectionString = GetConnectionString()
    Else
        conn.connectionString = txtConnectionString.Text
    End If
    
    conn.Open
    
    insertSheetIndex = CInt(cboSheet.List(cboSheet.ListIndex, 1))
    Set shtTemplate = ThisWorkbook.Sheets(insertSheetIndex)
    
    For index = 0 To Me.lstTables.ListCount - 1
        If lstTables.selected(index) Then
            '-- Get the importing table name
            TableName = lstTables.List(index, 0)
            clearExistedData = Me.chkClearExistedData.Value
            
            '-- Get the sheet which is used to store the table information
            If Me.optImportModeOverwrite.Value Then
                Set shtCurrent = GetSheetFromTableName(TableName)
                If shtCurrent Is Nothing Then
                    Set shtCurrent = CopyASheet(shtTemplate, , ThisWorkbook.Sheets(insertSheetIndex))
                    insertSheetIndex = shtCurrent.index
                    clearExistedData = True
                End If
            ElseIf Me.optImportModeAlwaysCreateSheets.Value Then
                insertSheetIndex = ThisWorkbook.Sheets.Count
                Set shtCurrent = CopyASheet(shtTemplate, , ThisWorkbook.Sheets(insertSheetIndex))
            Else
                Set shtCurrent = ThisWorkbook.Sheets(insertSheetIndex)
            End If
            
            '-- Set public variant
            If isSetPublicVarient = False Then
                Me.ImportProvider.GetOptions().LastAccessTableName = TableName
                isSetPublicVarient = True
            End If
            Set table = Me.ImportProvider.GetLogicalTable(conn, TableName)
            '-- Write to sheet
            shtCurrent.Select
            Call basTableSheet.SetTableInfoToWorksheet(shtCurrent, table, clearExistedData)
        End If
    Next

    '-- Close connection
    conn.Close
    Set conn = Nothing
    
    Call SaveImportOptions
    basToolbar.Command_SetSheetsName_Click True
    Exit Sub
Flag_Err:
    Set conn = Nothing

    Call MsgBoxEx_Error
End Sub

Private Sub UpdateTextBoxConnectionString()
    Me.txtConnectionString = GetConnectionString
End Sub

Private Sub Init()
    On Error GoTo F_Error
    
    Call InitForm
    Call InitConnectionPage
    Call InitOptionPage
    Call SetWizardStatus(PageConnectIndex)
    
    Exit Sub

F_Error:
     Call MsgBoxEx_Error
End Sub

Private Sub InitConnectionPage()
    '-- Active connection page
    Me.MultiPageMain.Value = 0
    
    '-- Active connection sub page
    If Me.ImportProvider.GetOptions().ConnectionMode = ConnectionModeDataSource Then
        Me.MultiPageConnection.Value = 0
        txtServer.SetFocus
    Else
        Me.MultiPageConnection.Value = 1
        txtConnectionString.SetFocus
    End If
    
    '-- Init provider list
    Dim providers() As String
    Dim index As Integer
    
    cboProvider.Clear
    providers = Me.ImportProvider.providers
    For index = LBound(providers) To UBound(providers)
        Call cboProvider.AddItem(providers(index))
    Next
    
    cboProvider.ListIndex = 0
    If Len(Me.ImportProvider.GetOptions().Provider) > 0 Then
        cboProvider.Text = Me.ImportProvider.GetOptions().Provider
    End If

    '-- fill text box values
    Me.labDataSourceHelp = Me.ImportProvider.GetOptions().DataSourceTip
    txtServer.Text = Me.ImportProvider.GetOptions().DataSource
    txtUser.Text = Me.ImportProvider.GetOptions().UserName
    txtPassword.Text = Me.ImportProvider.GetOptions().Password
    
    If Me.ImportProvider.GetOptions().ConnectionMode = ConnectionModeConnectionString Then
        txtConnectionString.Text = Me.ImportProvider.GetOptions().connectionString
    End If
    
    '-- Init Database
    Me.cboDatabase.Clear
    If Me.ImportProvider.SupportSelectDatabase Then
        Me.btnRefreshDatabase.Enabled = True
        Me.labDatabase.Enabled = True
        Me.cboDatabase.Enabled = True
    Else
        Me.btnRefreshDatabase.Enabled = False
        Me.labDatabase.Enabled = False
        Me.cboDatabase.Enabled = False
    End If
End Sub

Private Sub InitOptionPage()
    Dim iActiveSheet As Integer
    Dim iSheet As Integer
    Dim index As Integer
    Dim shtCurrent As Worksheet

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
    
    '-- init importing options
    Select Case Me.ImportProvider.GetOptions().ImportMode
    Case enmImportMode.ImportModeOverwrite
        Me.optImportModeOverwrite.Value = True
    Case enmImportMode.ImportModeAlwaysCreateSheet
        Me.optImportModeAlwaysCreateSheets.Value = True
    Case enmImportMode.ImportModeAlwaysUpdate
        Me.optImportModeOnlyUpdateTemplateSheet.Value = True
    End Select
    
    Me.chkClearExistedData.Value = Me.ImportProvider.GetOptions().ClearDataInExistedSheet
End Sub

Private Sub InitForm()
    Me.Caption = "Import from " & DatabaseType
End Sub

Private Sub SaveConnectionOptions()
    If Me.MultiPageConnection.Value = 0 Then
        Me.ImportProvider.GetOptions().ConnectionMode = ConnectionModeDataSource
    Else
        Me.ImportProvider.GetOptions().ConnectionMode = ConnectionModeConnectionString
    End If
    
    Me.ImportProvider.GetOptions().Provider = Trim(cboProvider.Text)
    Me.ImportProvider.GetOptions().DataSource = Trim(txtServer.Text)
    Me.ImportProvider.GetOptions().UserName = Trim(txtUser.Text)
    Me.ImportProvider.GetOptions().Password = txtPassword.Text
    Me.ImportProvider.GetOptions().LastDatabaseName = Me.cboDatabase.Text
    Me.ImportProvider.GetOptions().connectionString = Trim(txtConnectionString.Text)
End Sub

Private Sub SaveImportOptions()
    If Me.optImportModeOverwrite.Value Then
        Me.ImportProvider.GetOptions().ImportMode = ImportModeOverwrite
    ElseIf Me.optImportModeAlwaysCreateSheets.Value Then
        Me.ImportProvider.GetOptions().ImportMode = ImportModeAlwaysCreateSheet
    Else
        Me.ImportProvider.GetOptions().ImportMode = ImportModeAlwaysUpdate
    End If
    
    Me.ImportProvider.GetOptions().ClearDataInExistedSheet = Me.chkClearExistedData.Value
End Sub

Private Sub ShowTablesInfo()
    Dim selectTableCount As Integer
    Dim index As Integer
    
    selectTableCount = 0
    For index = 0 To Me.lstTables.ListCount - 1
        If lstTables.selected(index) Then
            selectTableCount = selectTableCount + 1
        End If
    Next
    Me.labTable.Caption = "Select Tables (" & CStr(selectTableCount) & "\" & CStr(Me.lstTables.ListCount) & ")"
End Sub

Private Sub btnConnBuild_Click()
    On Error GoTo Flag_Err
    
    Me.txtConnectionString.Text = basPublicDatabase.GetConnectionString(Me.txtConnectionString.Text)
    Exit Sub
    
Flag_Err:

    Call MsgBoxEx_Error
End Sub

Private Sub btnNext_Click()
    On Error GoTo Flag_Err
    
    Select Case Me.MultiPageMain.Value
    Case PageConnectIndex
        Call DoConnect
    Case PageTablesIndex
    
    Case PageOptionIndex
        Call DoImport
    End Select
    
    If Me.MultiPageMain.Value < PageOptionIndex Then
        Call SetWizardStatus(Me.MultiPageMain.Value + 1)
    End If
    
    Exit Sub
Flag_Err:
    Call MsgBoxEx_Error
End Sub

Private Sub btnPrevious_Click()
    On Error GoTo Flag_Err
    
    If Me.MultiPageMain.Value > 0 Then
        Me.MultiPageMain.Value = Me.MultiPageMain.Value - 1
        Call SetWizardStatus(Me.MultiPageMain.Value - 1)
    End If
    
    Exit Sub
Flag_Err:
    Call MsgBoxEx_Error
End Sub

Private Sub btnRefreshDatabase_Click()
    Dim conn As ADODB.Connection
    Dim sSQL As String
    Dim oRs As ADODB.Recordset
    Dim lastIndex As Integer
    Dim index As Integer

    On Error GoTo Flag_Err

    If FillDatabasesFromNames() Then
        Exit Sub
    End If
    
    Set conn = New ADODB.Connection
    conn.connectionString = GetConnectionString
    conn.Open
    sSQL = Me.ImportProvider.GetDatabasesSql()

    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    lastIndex = -1
    index = -1
    cboDatabase.Clear
    Do While Not oRs.EOF
        index = index + 1
        cboDatabase.AddItem (oRs("name"))
        If oRs("name") = Me.ImportProvider.GetOptions().LastDatabaseName Then
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

    Exit Sub
Flag_Err:
    Set oRs = Nothing
    Set conn = Nothing

    Call MsgBoxEx_Error
End Sub

Private Function FillDatabasesFromNames() As Boolean
    Dim lastIndex As Integer
    Dim index As Integer
    Dim objDatabases As Collection
    
    Set objDatabases = Me.ImportProvider.GetDatabases(Trim(txtServer.Text), _
                            Trim(txtUser.Text), _
                            Trim(txtPassword.Text))
    If objDatabases Is Nothing Then
        FillDatabasesFromNames = False
        Exit Function
    End If
    
    cboDatabase.Clear
    For index = 1 To objDatabases.Count
        cboDatabase.AddItem objDatabases.Item(index)
        If objDatabases.Item(index) = Me.ImportProvider.GetOptions().LastDatabaseName Then
            lastIndex = index - 1
        End If
    Next

    If cboDatabase.ListCount > 0 Then
        If lastIndex > 0 Then
            cboDatabase.ListIndex = lastIndex
        Else
            cboDatabase.ListIndex = 0
        End If
    End If
    
    FillDatabasesFromNames = True
End Function

Private Sub cboDatabase_Change()
    Call UpdateTextBoxConnectionString
End Sub

Private Sub cboProvider_Change()
    Call UpdateTextBoxConnectionString
End Sub

Private Sub cboSheet_Enter()
    Call SelectAllText(cboSheet)
End Sub

Private Sub chkSelectTablesAll_Change()

    SkipShowTablesStatisticsInfo = True
    Call SelectAllListBoxItems(Me.lstTables, chkSelectTablesAll.Value)
    SkipShowTablesStatisticsInfo = False
    lstTables_Change

End Sub

Private Sub lstTables_Change()
    If Not SkipShowTablesStatisticsInfo Then
        Call ShowTablesInfo
    End If
End Sub

Private Sub optImportModeOnlyUpdateTemplateSheet_Click()
    If optImportModeOnlyUpdateTemplateSheet.Value Then
        Me.chkClearExistedData.Value = True
    End If
End Sub

Private Sub optImportModeOverwrite_Change()
    If optImportModeOverwrite.Value Then
        Me.chkClearExistedData.Value = True
    End If
End Sub

Private Sub txtPassword_Change()
    Call UpdateTextBoxConnectionString
End Sub

Private Sub txtServer_Change()
    Call UpdateTextBoxConnectionString
End Sub

Private Sub txtServer_Enter()
    Call SelectAllText(txtServer)
End Sub

Private Sub txtPassword_Enter()
    Call SelectAllText(txtPassword)
End Sub

Private Sub txtUser_Change()
    Call UpdateTextBoxConnectionString
End Sub

Private Sub txtUser_Enter()
    Call SelectAllText(txtUser)
End Sub

Private Sub UserForm_Activate()
    If mInitialized Then Exit Sub
    
    SkipShowTablesStatisticsInfo = False
    Call Init
    mInitialized = True
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

