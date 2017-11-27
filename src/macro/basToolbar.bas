Attribute VB_Name = "basToolbar"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2013, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit
Public Const BAR_NAME                   As String = "Database Modeling Bar"
Public Const BAR_BUTTON_ID              As Integer = 10000

Public barDBModeling                    As CommandBar
Public oMenuInfoCollection              As Collection

Private Sub InitMenuCollection()
    If Not oMenuInfoCollection Is Nothing Then
        Exit Sub
    End If
    
    Dim barIndex     As Integer
    barIndex = 0
    
    Set oMenuInfoCollection = New Collection
    '-- Generate Menu
    Select Case The_Excel_Type
        Case DBName_All
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
                msoControlPopup, _
                "Generate", _
                "Generate script", _
                "", _
                "", _
                barIndex), "Generate")
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
                msoControlButton, _
                "SQL Server...", _
                "", _
                "Command_GenerateSQLServer_Click", _
                "Generate", _
                barIndex), "SQL Server")
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
                msoControlButton, _
                "DB2...", _
                "", _
                "Command_GenerateDB2_Click", _
                "Generate", _
                barIndex), "DB2")
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
                msoControlButton, _
                "MariaDB...", _
                "", _
                "Command_GenerateMariaDB_Click", _
                "Generate", _
                barIndex), "MariaDB")
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
                msoControlButton, _
                "MySQL...", _
                "", _
                "Command_GenerateMySQL_Click", _
                "Generate", _
                barIndex), "MySQL")
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
                msoControlButton, _
                "Oracle...", _
                "", _
                "Command_GenerateOracle_Click", _
                "Generate", _
                barIndex), "Oracle")
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
                msoControlButton, _
                "PostgreSQL...", _
                "", _
                "Command_GeneratePostgreSQL_Click", _
                "Generate", _
                barIndex), "PostgreSQL")
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
                msoControlButton, _
                "SQLite...", _
                "", _
                "Command_GenerateSQLite_Click", _
                "Generate", _
                barIndex), "SQLite")
    Case DBName_SQLServer
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
            msoControlButton, _
            "Generate", _
            "Generate script", _
            "Command_GenerateSQLServer_Click", _
            "", _
            barIndex), "Generate")
    Case DBName_DB2
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
            msoControlButton, _
            "Generate", _
            "Generate script", _
            "Command_GenerateDB2_Click", _
            "", _
            barIndex), "Generate")
    Case DBName_MariaDB
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
            msoControlButton, _
            "Generate", _
            "Generate script", _
            "Command_GenerateMariaDB_Click", _
            "", _
            barIndex), "Generate")
    Case DBName_MySQL
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
            msoControlButton, _
            "Generate", _
            "Generate script", _
            "Command_GenerateMySQL_Click", _
            "", _
            barIndex), "Generate")
    Case DBName_Oracle
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
            msoControlButton, _
            "Generate", _
            "Generate script", _
            "Command_GenerateOracle_Click", _
            "", _
            barIndex), "Generate")
    Case DBName_PostgreSQL
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
            msoControlButton, _
            "Generate", _
            "Generate script", _
            "Command_GeneratePostgreSQL_Click", _
            "", _
            barIndex), "Generate")
    Case DBName_SQLite
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
            msoControlButton, _
            "Generate", _
            "Generate script", _
            "Command_GenerateSQLite_Click", _
            "", _
            barIndex), "Generate")
    End Select
    
    '-- Tools menu
    Call oMenuInfoCollection.Add(GetMenuInfoObject( _
            msoControlPopup, _
            "Tools", _
            "Tools", _
            "", _
            "", _
            barIndex), "DME.Tools")
    Call oMenuInfoCollection.Add(GetMenuInfoObject( _
            msoControlButton, _
            "Add Hyperlinks in The Index Sheet", _
            "Add Hyperlinks in The Index Sheet", _
            "Command_AddHyperlinks_Click", _
            "DME.Tools", _
            barIndex), "Add Hyperlinks in The Index Sheet")
    Call oMenuInfoCollection.Add(GetMenuInfoObject( _
            msoControlButton, _
            "Set Sheets Name", _
            "Set sheets name like <Table Name>", _
            "Command_SetSheetsName_Click", _
            "DME.Tools", _
            barIndex), "Set Sheets Name")
    Call oMenuInfoCollection.Add(GetMenuInfoObject( _
            msoControlButton, _
            "Sort Sheets by Sheet Names", _
            "Sort Sheets by Sheet Names", _
            "Command_SortSheetByName_Click", _
            "DME.Tools", _
            barIndex), "Sort Sheets by Sheet Names")

    '-- Import menu
    Select Case The_Excel_Type
        Case DBName_All
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
                msoControlPopup, _
                "Import", _
                "Import database to excel sheet.", _
                "", _
                "", _
                barIndex), "Import")
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
                msoControlButton, _
                "SQL Server...", _
                "", _
                "Command_Import_SQLServer_Click", _
                "Import", _
                barIndex), "Import_SQLServer")
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
                msoControlButton, _
                "DB2...", _
                "", _
                "Command_Import_DB2_Click", _
                "Import", _
                barIndex), "Import_DB2")
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
                msoControlButton, _
                "MariaDB...", _
                "", _
                "Command_Import_MariaDB_Click", _
                "Import", _
                barIndex), "Import_MariaDB")
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
                msoControlButton, _
                "MySQL...", _
                "", _
                "Command_Import_MySQL_Click", _
                "Import", _
                barIndex), "Import_MySQL")
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
                msoControlButton, _
                "Oracle...", _
                "", _
                "Command_Import_Oracle_Click", _
                "Import", _
                barIndex), "Import_Oracle")
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
                msoControlButton, _
                "PostgreSQL...", _
                "", _
                "Command_Import_PostgreSQL_Click", _
                "Import", _
                barIndex), "Import_PostgreSQL")
    Case DBName_SQLServer
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
            msoControlButton, _
            "Import", _
            "Import database to excel sheet.", _
            "Command_Import_SQLServer_Click", _
            "", _
            barIndex), "Import")
    Case DBName_DB2
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
            msoControlButton, _
            "Import", _
            "Import database to excel sheet.", _
            "Command_Import_DB2_Click", _
            "", _
            barIndex), "Import")
    Case DBName_MariaDB
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
            msoControlButton, _
            "Import", _
            "Import database to excel sheet.", _
            "Command_Import_MariaDB_Click", _
            "", _
            barIndex), "Import")
    Case DBName_MySQL
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
            msoControlButton, _
            "Import", _
            "Import database to excel sheet.", _
            "Command_Import_MySQL_Click", _
            "", _
            barIndex), "Import")
    Case DBName_Oracle
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
            msoControlButton, _
            "Import", _
            "Import database to excel sheet.", _
            "Command_Import_Oracle_Click", _
            "", _
            barIndex), "Import")
    Case DBName_PostgreSQL
        Call oMenuInfoCollection.Add(GetMenuInfoObject( _
            msoControlButton, _
            "Import", _
            "Import database to excel sheet.", _
            "Command_Import_PostgreSQL_Click", _
            "", _
            barIndex), "Import")
    End Select

    Call oMenuInfoCollection.Add(GetMenuInfoObject( _
            msoControlButton, _
            "About...", _
            "", _
            "Command_About_Click", _
            "", _
            barIndex), "About")
End Sub

Private Function GetMenuInfoObject(ByVal Style As MsoControlType _
                    , ByVal Caption As String _
                    , ByVal TooltipText As String _
                    , ByVal OnAction As String _
                    , ByVal Parent As String _
                    , ByRef barIndex As Integer) As clsMenuInfo
    Dim oMenu As clsMenuInfo
    Dim oParentMenu  As clsMenuInfo
    
    Set oMenu = New clsMenuInfo
    oMenu.Style = Style
    oMenu.Caption = Caption
    oMenu.TooltipText = TooltipText
    oMenu.OnAction = OnAction
    oMenu.Parent = Parent
    oMenu.ChildCount = 0
    oMenu.InstanceIndex = 0
    Set oMenu.Instance = Nothing
    
    '-- set instance index
    If Len(oMenu.Parent) = 0 Then
        barIndex = barIndex + 1
        oMenu.InstanceIndex = barIndex
    Else
        Set oParentMenu = oMenuInfoCollection.Item(oMenu.Parent)
        
        oParentMenu.ChildCount = oParentMenu.ChildCount + 1
        oMenu.InstanceIndex = oParentMenu.ChildCount
    End If
    
    '-- return
    Set GetMenuInfoObject = oMenu
End Function

Private Sub CreateCommandBarButtons()
    Call InitMenuCollection
    
    Dim oMenu        As clsMenuInfo
    Dim oParentMenu  As clsMenuInfo
    Dim oParentPopup As CommandBarPopup
    Dim oButton      As CommandBarButton
    Dim oPopup       As CommandBarPopup
    Dim iFaceID      As Integer
    iFaceID = BAR_BUTTON_ID
    
    For Each oMenu In oMenuInfoCollection
        If Len(oMenu.Parent) > 0 Then
            Set oParentMenu = oMenuInfoCollection.Item(oMenu.Parent)
            
            Set oParentPopup = oParentMenu.Instance
            If oMenu.Style = msoControlPopup Then
                Set oPopup = oParentPopup.Controls.Add(msoControlPopup, , , , True)
            Else
                Set oButton = oParentPopup.Controls.Add(msoControlButton, , , , True)
            End If
        Else '-- first level menu items
            
            If oMenu.Style = msoControlPopup Then
                Set oPopup = barDBModeling.Controls.Add(msoControlPopup, , , , True)
            Else
                Set oButton = barDBModeling.Controls.Add(msoControlButton, , , , True)
            End If
        End If
        
        If oMenu.Style = msoControlPopup Then
            With oPopup
                .Caption = oMenu.Caption
                .TooltipText = oMenu.TooltipText
                .OnAction = ""
                .Visible = True
                .BeginGroup = True
            End With
            Set oMenu.Instance = oPopup
        Else
            With oButton
                .Style = msoButtonCaption
                .Caption = oMenu.Caption
                .TooltipText = oMenu.TooltipText
                .OnAction = oMenu.OnAction
                .Visible = True
                .BeginGroup = True
                '- .faceID = iFaceID
            End With
            Set oMenu.Instance = oButton
        End If
    Next
    
    Set oMenu = Nothing
    Set oParentMenu = Nothing
    Set oParentPopup = Nothing
    Set oButton = Nothing
    Set oPopup = Nothing
End Sub

Private Sub ResetCommandBarButtonsOnAction()
    Dim oMenu As clsMenuInfo
    
    For Each oMenu In oMenuInfoCollection
        If oMenu.Style = msoControlPopup Then
        Else
            oMenu.Instance.OnAction = oMenu.OnAction
        End If
    Next
    
    Set oMenu = Nothing
End Sub

Public Sub SetCommandBarButtonsToNothing()
    Dim oMenu As clsMenuInfo
    If oMenuInfoCollection Is Nothing Then Exit Sub
    For Each oMenu In oMenuInfoCollection
        Set oMenu.Instance = Nothing
    Next
    
    Set oMenu = Nothing
End Sub

Private Sub GetCommandBarButtonsInstance()
    Dim oMenu        As clsMenuInfo
    Dim oParentMenu  As clsMenuInfo
    
    Call InitMenuCollection
    
    For Each oMenu In oMenuInfoCollection
        If Len(oMenu.Parent) > 0 Then
            Set oParentMenu = oMenuInfoCollection.Item(oMenu.Parent)

            Set oMenu.Instance = oParentMenu.Instance.Controls(oMenu.InstanceIndex)
        Else '-- first level menu items
            Set oMenu.Instance = barDBModeling.Controls(oMenu.InstanceIndex)
        End If
    Next
    
    Set oMenu = Nothing
    Set oParentMenu = Nothing
End Sub

'--------------------------------------
'() add this command bar
'--------------------------------------
Public Sub AddCommandBar()
    Call DeleteCommandBar
    
    On Error Resume Next
    '-- if another excel is already add the bar, we just use it
    Set barDBModeling = CommandBars.Item(BAR_NAME)
    If Err.Number = 0 Then
        If barDBModeling.Controls.Count >= 4 Then
            Set barDBModeling = CommandBars.Item(BAR_NAME)
            
            GetCommandBarButtonsInstance
        Else
            Set barDBModeling = Nothing
        End If
    Else
        Set barDBModeling = Nothing
    End If
    
    On Error GoTo 0
    '-- add a new bar
    If barDBModeling Is Nothing Then
        Set barDBModeling = CommandBars.Add(, , , True)
        On Error Resume Next
        With barDBModeling
            .Name = BAR_NAME
            
            Call CreateCommandBarButtons

            .Position = msoBarTop
            .Visible = True
        End With
    Else
        barDBModeling.Visible = True
        Call ResetCommandBarButtonsOnAction
    End If
End Sub

Public Sub DeleteCommandBar()
    On Error Resume Next
    Set barDBModeling = CommandBars.Item(BAR_NAME)
    If Not barDBModeling Is Nothing Then
        barDBModeling.Visible = False
    End If
End Sub

Public Sub Command_AddHyperlinks_Click()
   frmAddHyperlinks.Show
End Sub

Public Sub Command_GenerateSQLServer_Click()
    frmGenerateSQL.DatabaseType = DBName_SQLServer
    frmGenerateSQL.Show
End Sub

Public Sub Command_GenerateDB2_Click()
    frmGenerateSQL.DatabaseType = DBName_DB2
    frmGenerateSQL.Show
End Sub

Public Sub Command_GenerateMariaDB_Click()
    frmGenerateSQL.DatabaseType = DBName_MariaDB
    frmGenerateSQL.Show
End Sub

Public Sub Command_GenerateMySQL_Click()
    frmGenerateSQL.DatabaseType = DBName_MySQL
    frmGenerateSQL.Show
End Sub

Public Sub Command_GenerateOracle_Click()
    frmGenerateSQL.DatabaseType = DBName_Oracle
    frmGenerateSQL.Show
End Sub

Public Sub Command_GeneratePostgreSQL_Click()
    frmGenerateSQL.DatabaseType = DBName_PostgreSQL
    frmGenerateSQL.Show
End Sub

Public Sub Command_GenerateSQLite_Click()
    frmGenerateSQL.DatabaseType = DBName_SQLite
    frmGenerateSQL.Show
End Sub

Public Sub Command_Import_SQLServer_Click()
    frmImport.DatabaseType = DBName_SQLServer
    frmImport.Show
End Sub

Public Sub Command_Import_DB2_Click()
    frmImport.DatabaseType = DBName_DB2
    frmImport.Show
End Sub

Public Sub Command_Import_MariaDB_Click()
    frmImport.DatabaseType = DBName_MariaDB
    frmImport.Show
End Sub

Public Sub Command_Import_MySQL_Click()
    frmImport_MySQL.Show
End Sub

Public Sub Command_Import_Oracle_Click()
    frmImport.DatabaseType = DBName_Oracle
    frmImport.Show
End Sub

Public Sub Command_Import_PostgreSQL_Click()
    frmImport.DatabaseType = DBName_PostgreSQL
    frmImport.Show
End Sub

'---------------------
'-  Sort Sheet By Table Name
'---------------------
Public Sub Command_SortSheetByName_Click()
    Dim iSheet          As Integer
    Dim iSheet2         As Integer
    Dim TableName       As String
    Dim tableName2      As String
    Dim sh              As Worksheet
    Dim sh2             As Worksheet
    Dim iCurSheet       As Integer
    
    On Error GoTo Flag_Err
    iCurSheet = ThisWorkbook.ActiveSheet.index
    
    For iSheet = Sheet_First_Table To ThisWorkbook.Sheets.Count
        Set sh = ThisWorkbook.Sheets(iSheet)
        TableName = Trim(sh.Cells(Table_Sheet_Row_TableName, Table_Sheet_Col_TableName).Value)
        For iSheet2 = iSheet + 1 To ThisWorkbook.Sheets.Count
            Set sh2 = ThisWorkbook.Sheets(iSheet2)
            tableName2 = Trim(sh2.Cells(Table_Sheet_Row_TableName, Table_Sheet_Col_TableName).Value)
            
            If VBA.StrComp(TableName, tableName2, vbBinaryCompare) > 0 Then
                Call sh2.Move(Before:=sh)
                sh2.Name = sh2.Name
                sh.Name = sh.Name
                Set sh = ThisWorkbook.Sheets(iSheet)
                TableName = Trim(sh.Cells(Table_Sheet_Row_TableName, Table_Sheet_Col_TableName).Value)
            End If
        Next
    Next
    
    On Error Resume Next
    '-- re-calculating first sheet
    Set sh = ThisWorkbook.Sheets(Sheet_Index)
    sh.Activate
    sh.EnableCalculation = False
    sh.EnableCalculation = True
    
    Set sh = ThisWorkbook.Sheets(iCurSheet)
    sh.Activate
    Exit Sub
    
Flag_Err:
    MsgBox Err.Description, vbInformation, App_Name
End Sub

'---------------------
'-  Set Sheet Name base one Sheet Information
'---------------------
Public Sub Command_SetSheetsName_Click(Optional ignoreError As Boolean = False)
    Dim iSheet          As Integer
    Dim index           As Integer
    Dim shtCurrent      As Worksheet
    Dim sTableCaption   As String
    Dim sheetName       As String
    Dim iRow            As Integer
    Dim iCol            As String
    
    If ignoreError Then
        On Error Resume Next
    Else
        On Error GoTo Flag_Err
    End If
    If Sheet_NameIsTableDesc Then
        iRow = Table_Sheet_Row_TableComment
        iCol = Table_Sheet_Col_TableComment
    Else
        iRow = Table_Sheet_Row_TableName
        iCol = Table_Sheet_Col_TableName
    End If
    
    index = 0
    For iSheet = Sheet_First_Table To ThisWorkbook.Sheets.Count
        index = index + 1
        Set shtCurrent = ThisWorkbook.Sheets(iSheet)
        
        If Table_Code_Length = 0 Then
            '-- Just name
            sheetName = Trim(shtCurrent.Cells.Item(iRow, iCol).Text)
        Else
            '-- like 000.Employee
            sheetName = Format(index, String(Table_Code_Length, "0")) & "." _
                        & Trim(shtCurrent.Cells.Item(iRow, iCol).Text)
        End If
        
        If Len(sheetName) > 31 Then
            sheetName = Left(sheetName, 31)
        End If
        shtCurrent.Name = sheetName
    Next
    
    '-- re-calculating first sheet
    Set shtCurrent = ThisWorkbook.Sheets(Sheet_Index)
    shtCurrent.EnableCalculation = False
    shtCurrent.EnableCalculation = True
    Exit Sub
    
Flag_Err:
    MsgBox "An error occurs when set a sheet name with " & sheetName _
            & Line & Err.Description, _
            vbInformation, App_Name
End Sub

'---------------------
'-  About
'---------------------
Public Sub Command_About_Click()
    frmAbout.Show
End Sub

'---------------------
'-  Get Table Name
'---------------------
Public Function GetTableName(index As Integer) As String
    On Error GoTo Flag_Err
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(index)
    
    If index < Sheet_First_Table Then
        GetTableName = sh.Name
    ElseIf Len(Trim(sh.Cells(Table_Sheet_Row_TableName, Table_Sheet_Col_TableName).Value)) > 0 Then
        GetTableName = Trim(sh.Cells(Table_Sheet_Row_TableName, Table_Sheet_Col_TableName).Value)
    Else
        If Table_Code_Length = 0 Then
            GetTableName = sh.Name
        Else
            GetTableName = Mid(sh.Name, Table_Code_Length + 2)
        End If
    End If
    
    Exit Function
Flag_Err:
    GetTableName = ""
End Function

'---------------------
'-  Get Table Name
'---------------------
Public Function GetTableNameFromSheet(sh As Worksheet) As String
    On Error GoTo Flag_Err
    
    If sh.index < Sheet_First_Table Then
        GetTableNameFromSheet = sh.Name
    ElseIf Len(Trim(sh.Cells(Table_Sheet_Row_TableName, Table_Sheet_Col_TableName).Value)) > 0 Then
        GetTableNameFromSheet = Trim(sh.Cells(Table_Sheet_Row_TableName, Table_Sheet_Col_TableName).Value)
    Else
        If Table_Code_Length = 0 Then
            GetTableNameFromSheet = sh.Name
        Else
            GetTableNameFromSheet = Mid(sh.Name, Table_Code_Length + 2)
        End If
    End If
    
    Exit Function
Flag_Err:
    GetTableNameFromSheet = ""
End Function

'---------------------
'-  Set hyperlinks
'---------------------
Public Function SetHyperlinks(SheetIndex As Integer, iRow As Integer, iCol As Integer) As String
    On Error Resume Next
    Dim strText     As String
    Dim objCell     As Range
    Dim strSheetName    As String
    On Error Resume Next
    strSheetName = GetSheetName(SheetIndex)
    
    If Len(Trim(strSheetName)) = 0 Then
        SetHyperlinks = ""
        
    Else
        SetHyperlinks = ">>"
        
        Set objCell = ThisWorkbook.Sheets(Sheet_Index).Cells(iRow, iCol)
        objCell.Hyperlinks.Delete
        Call objCell.Hyperlinks.Add(objCell, "")
        objCell.Hyperlinks(1).SubAddress = "'" & strSheetName & "'!A1"
    End If
End Function


