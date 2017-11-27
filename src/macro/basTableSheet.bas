Attribute VB_Name = "basTableSheet"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2012, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

'---------------------------------------------
'-  Get All Table information
'---------------------------------------------
Public Function GetAllLogicalTables(Optional includeIgnoreTables As Boolean = False) As Collection
    Dim objLogicalTables    As Collection
    Dim iSheet              As Integer
    Dim oSheet              As Worksheet
    
    Set objLogicalTables = New Collection
    
    For iSheet = Sheet_First_Table To ThisWorkbook.Sheets.Count
        Set oSheet = ThisWorkbook.Sheets(iSheet)
        If includeIgnoreTables _
            Or VBA.StrComp( _
                LCase(TrimEx( _
                    oSheet.Cells.Item(Table_Sheet_Row_TableStatus, Table_Sheet_Col_TableStatus).Text)) _
                , Table_Sheet_TableStatus_Ignore) _
            <> 0 Then
            objLogicalTables.Add GetTableInfoFromWorksheet(ThisWorkbook.Sheets(iSheet))
        End If
    Next
    
    '-- Return
    Set GetAllLogicalTables = objLogicalTables
End Function

'---------------------------------------------
'-  Get Table information
'---------------------------------------------
Public Function GetTableInfoFromWorksheet(shtCurrent As Worksheet) As clsLogicalTable
    On Error GoTo Flag_Err
    
    Dim contextSheet As String
    Dim contextArea As String
    
    contextSheet = shtCurrent.Name
    Dim objTable As clsLogicalTable
    
    Set objTable = New clsLogicalTable
    contextArea = "Table Name"
    objTable.TableName = Trim(shtCurrent.Cells.Item(Table_Sheet_Row_TableName, Table_Sheet_Col_TableName).Text)
    
    contextArea = "Table Comment"
    objTable.Comment = Trim(shtCurrent.Cells.Item(Table_Sheet_Row_TableComment, Table_Sheet_Col_TableComment).Text)
   
    contextArea = "Primary Key"
    Set objTable.PrimaryKey = GetTablePrimaryKey(shtCurrent)
    
    contextArea = "Foreign Keys"
    Set objTable.ForeignKeys = GetTableForeignKeys(shtCurrent, objTable)
    
    contextArea = "Indexes"
    Set objTable.Indexes = GetTableIndexes(shtCurrent, objTable)
    
    contextArea = "Columns"
    Set objTable.Columns = GetTableColumns(shtCurrent)
    
    '-- Return
    Set GetTableInfoFromWorksheet = objTable
    Exit Function
    
Flag_Err:
    Dim errorMessage As String
    errorMessage = Err.Description _
                & Line & "Sheet Name: " & contextSheet _
                & Line & "Area: " & contextArea
    Err.Raise 1, , errorMessage
End Function

'---------------------------------------------
'-  Get Columns information
'---------------------------------------------
Public Function GetTableColumns(shtCurrent As Worksheet) As Collection
    Dim objColumns      As Collection
    Dim objColumn       As clsLogicalColumn
    Dim strCell         As String
    Dim index           As Integer
    
    Set objColumns = New Collection
    index = 1
    Do While (True)
        '-- Get Column name
        strCell = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_First_Column + index - 1, Table_Sheet_Col_ColumnName).Text)
        
        '-- if Column name is '', finished Columns search
        If Len(strCell) = 0 Then Exit Do
        
        '-- add a Column information
        Set objColumn = New clsLogicalColumn
        objColumns.Add objColumn
        With objColumn
            '-- Get Column information
            .ColumnLabel = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_First_Column + index - 1, Table_Sheet_Col_ColumnLabel).Text)
            .ColumnName = strCell
            .DataType = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_First_Column + index - 1, Table_Sheet_Col_ColumnDataType).Text)
            .Nullable = IIf(UCase(Trim(shtCurrent.Cells.Item(Table_Sheet_Row_First_Column + index - 1, Table_Sheet_Col_ColumnNullable).Text)) = "YES", True, False)
            .Default = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_First_Column + index - 1, Table_Sheet_Col_ColumnDefault).Text)
            .Comment = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_First_Column + index - 1, Table_Sheet_Col_ColumnComment).Text)
        End With
        index = index + 1
    Loop
    
    '-- Return
    Set GetTableColumns = objColumns
End Function

'---------------------------------------------
'-  Get PrimaryKeys information
'---------------------------------------------
Public Function GetTablePrimaryKey(shtCurrent As Worksheet) As clsLogicalPrimaryKey
    Dim objPK           As clsLogicalPrimaryKey
    Dim strCell         As String
    
    Set objPK = New clsLogicalPrimaryKey
    
    strCell = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_PrimaryKey, Table_Sheet_Col_Clustered).Text)
    With objPK
        '-- Get PK Columns' information
        .PKcolumns = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_PrimaryKey, Table_Sheet_Col_PrimaryKey).Text)
        
        '-- Get clustered information
        .IsClustered = Not (UCase(Trim(strCell)) = "N")
    End With
    
    '-- Return
    Set GetTablePrimaryKey = objPK
End Function

'---------------------------------------------
'-  Get ForeignKeys information
'---------------------------------------------
Public Function GetTableForeignKeys(shtCurrent As Worksheet, _
                                objTable As clsLogicalTable) As Collection
    Dim objFKs          As Collection
    Dim objFk           As clsLogicalForeignKey
    Dim strCell         As String
    Dim index           As Integer
    Dim strFKArray()    As String
    Dim strFKItem       As String
    Dim intFKItemLen    As Integer
    Dim intPos          As Integer
    
    index = 0
    Set objFKs = New Collection
    
    '-- Get Column name
    strCell = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_ForeignKey, Table_Sheet_Col_ForeignKey).Text)
    '-- Split PK infomation into array, one item is a infomation of foreign key
    strFKArray = Split(strCell, ";")
    
    For index = 0 To UBound(strFKArray)
        strFKItem = TrimEx(strFKArray(index))
        If Len(strFKItem) = 0 Then
            GoTo Flag_next
        End If
        
        Set objFk = New clsLogicalForeignKey
        objFKs.Add objFk
        
        '-- Replace ", "  to ",", to aviod get wrong table name.
        Do While True
            intFKItemLen = Len(strFKItem)
            strFKItem = Replace(strFKItem, ", ", ",")
            If Len(strFKItem) = intFKItemLen Then
                Exit Do
            End If
        Loop
        
        intPos = InStr(1, strFKItem, " ")
        With objFKs(index + 1)
            '-- Get Foreign key's Columnname
            .FKcolumns = Left(strFKItem, intPos - 1)
            '-- get rid of vbcr and vblf
            Do While Left(.FKcolumns, 1) = vbCr Or Left(.FKcolumns, 1) = vbLf
                .FKcolumns = Mid(.FKcolumns, 2)
            Loop
            '-- Get Foreign key's foreign table infomation
            Call SetForeignKeyRefTableAndName(objFKs(index + 1), Mid(strFKItem, intPos + 1))
            
            .FKName = Replace(Replace(.FKcolumns, " ", ""), ",", "$")
        End With
Flag_next:
    Next
    
    '-- Return
    Set GetTableForeignKeys = objFKs
End Function

'---------------------------------------------
'-  GetIndexKeys information
'---------------------------------------------
Public Function GetTableIndexes(shtCurrent As Worksheet, _
                                objTable As clsLogicalTable) As Collection
    Dim objIKs              As Collection
    Dim objIK               As clsLogicalIndex
    Dim strCell             As String
    Dim index               As Integer
    Dim strIKArray()        As String
    Dim strIKUnique()       As String
    Dim strIKClustered()    As String
    Dim strIKItem           As String
    Dim intPos              As Integer
    
    index = 0
    Set objIKs = New Collection
  
    '-- Get Index infomation
    strCell = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_Index, Table_Sheet_Col_Index).Text)
    '-- Split index infomation into array, one item is a infomation of index
    strIKArray = Split(strCell, ";")
    
    '-- Get index Unique information
    strCell = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_Index, Table_Sheet_Col_Unique).Text)
    '-- Split index's unique infomation into array, one item is a infomation of index's unique
    strIKUnique = Split(strCell, ";")
    
    '-- Get index Clustered information
    strCell = TrimEx(shtCurrent.Cells.Item(Table_Sheet_Row_Index, Table_Sheet_Col_Clustered).Text)
    '-- Split index's Clustered infomation into array, one item is a infomation of index's Clustered
    strIKClustered = Split(strCell, ";")
    
    For index = 0 To UBound(strIKArray)
        '-- Get one IK describation
        strIKItem = TrimEx(strIKArray(index))
        If Len(strIKItem) = 0 Then
            GoTo Flag_next
        End If
        Set objIK = New clsLogicalIndex
        objIKs.Add objIK
        With objIK
            '-- Set default information
            .IsClustered = False
            .IsUnique = True
            
            '-- Get index's name and index's columns
            intPos = InStr(1, strIKItem, ",")
            If intPos = 0 Then
                '-- Get index's name
                .IKName = strIKItem
            Else
                '-- Get index's name
                .IKName = Replace(Replace(strIKItem, " ", ""), ",", "$")
            End If
            '-- Get index's Columns infomation
            .IKColumns = "(" & strIKItem & ")"
            
            '-- Is Uniqued?
            If UBound(strIKUnique) >= index Then
                If UCase(TrimEx(strIKUnique(index))) = "N" Then
                    '-- Not Unique flag
                    .IsUnique = False
                End If
            End If
            
            '-- Is Clustered?
            If UBound(strIKClustered) >= index Then
                If UCase(TrimEx(strIKClustered(index))) = "Y" Then
                    '-- Not Unique flag
                    .IsClustered = True
                End If
            End If
        End With
Flag_next:
    Next
    
    '-- Return
    Set GetTableIndexes = objIKs
End Function

'---------------------------------------------
'-  Write Table information to worksheet
'---------------------------------------------
Public Sub SetTableInfoToWorksheet(ByRef sh As Worksheet, _
                        ByRef table As clsLogicalTable, _
                        ByRef clearExistedData As Boolean)
    
    Dim indexText       As String
    Dim indexClustered  As String
    Dim indexUnique     As String
    Dim RowHeight       As Double
    
    RowHeight = sh.Rows("2:2").RowHeight
    '-- Set Table Name
    sh.Cells(Table_Sheet_Row_TableName, Table_Sheet_Col_TableName).Value = table.TableName
    
    If clearExistedData _
        Or sh.Cells(Table_Sheet_Row_TableComment, Table_Sheet_Col_TableComment).Text = "" Then
        sh.Cells.Item(Table_Sheet_Row_TableComment, Table_Sheet_Col_TableComment).Value = IIf(Len(table.Comment) > 0, "'" & table.Comment, "")
    End If
    
    '-- Set PK
    Call table.GetPrimaryKeyInfoText(indexText, indexClustered)
    sh.Cells(Table_Sheet_Row_PrimaryKey, Table_Sheet_Col_PrimaryKey).Value = indexText
    sh.Cells(Table_Sheet_Row_PrimaryKey, Table_Sheet_Col_Clustered).Value = indexClustered
    sh.Cells(Table_Sheet_Row_PrimaryKey, Table_Sheet_Col_Unique).Value = ""
    
    '-- Set Index
    Call table.GetIndexesInfoText(indexText, indexClustered, indexUnique)
    sh.Cells(Table_Sheet_Row_Index, Table_Sheet_Col_Index).Value = indexText
    sh.Cells(Table_Sheet_Row_Index, Table_Sheet_Col_Clustered).Value = indexClustered
    sh.Cells(Table_Sheet_Row_Index, Table_Sheet_Col_Unique).Value = indexUnique

    '-- Set Index Row Height
    sh.Rows(Table_Sheet_Row_Index & ":" & Table_Sheet_Row_Index).Select
    If table.Indexes.Count > 0 Then
        Selection.RowHeight = table.Indexes.Count * RowHeight
    Else
        Selection.RowHeight = 1 * RowHeight
    End If
    Application.CutCopyMode = False
    
    '-- Set FK
    sh.Cells(Table_Sheet_Row_ForeignKey, Table_Sheet_Col_ForeignKey).Value = table.GetForeignKeysText
    
    '-- Set FK Row Height
    sh.Rows(Table_Sheet_Row_ForeignKey & ":" & Table_Sheet_Row_ForeignKey).Select
    If table.ForeignKeys.Count > 0 Then
        Selection.RowHeight = table.ForeignKeys.Count * RowHeight
    Else
        Selection.RowHeight = 1 * RowHeight
    End If
    Application.CutCopyMode = False
    
    '-- Set Column
    Dim Row                 As Integer
    Dim tableColumn         As clsLogicalColumn
    
    Row = Table_Sheet_Row_First_Column
    For Each tableColumn In table.Columns
        '-- Render column row
        If Row > Table_Sheet_Row_First_Column And sh.Cells(Row, Table_Sheet_Col_ColumnName).Text = "" Then
            Call InsertColumnRow(sh, Row)
        End If
        sh.Range(Table_Sheet_Col_ColumnLabel & Row, GetPreviousColumnName(Table_Sheet_Col_ColumnDataType) & Row).Interior.ColorIndex = xlNone
        sh.Range(Table_Sheet_Col_ColumnLabel & Row, GetPreviousColumnName(Table_Sheet_Col_ColumnDataType) & Row).Font.Bold = False
        '-- render PK
        If table.IsPKColumn(tableColumn.ColumnName) Then
            sh.Range(Table_Sheet_Col_ColumnLabel & Row, GetPreviousColumnName(Table_Sheet_Col_ColumnDataType) & Row).Interior.ColorIndex = 15
        End If
        '-- render FK
        If table.IsFKColumn(tableColumn.ColumnName) Then
            sh.Range(Table_Sheet_Col_ColumnLabel & Row, GetPreviousColumnName(Table_Sheet_Col_ColumnDataType) & Row).Font.Bold = True
        End If
        
        '-- set Column
        sh.Cells(Row, Table_Sheet_Col_ColumnName).Select
        If clearExistedData _
            Or sh.Cells(Row, Table_Sheet_Col_ColumnLabel).Text = "" Then
            
            sh.Cells(Row, Table_Sheet_Col_ColumnLabel).Value = tableColumn.ColumnLabel
        End If
        sh.Cells(Row, Table_Sheet_Col_ColumnName).Value = tableColumn.ColumnName
        sh.Cells(Row, Table_Sheet_Col_ColumnDataType).Value = tableColumn.DataType
        sh.Cells(Row, Table_Sheet_Col_ColumnNullable).Value = IIf(tableColumn.Nullable, Table_Sheet_Nullable, Table_Sheet_NonNullable)
        sh.Cells(Row, Table_Sheet_Col_ColumnDefault).Value = IIf(Len(tableColumn.Default) > 0, "'" & tableColumn.Default, "")
        If clearExistedData _
            Or sh.Cells(Row, Table_Sheet_Col_ColumnComment).Text = "" Then
            sh.Cells(Row, Table_Sheet_Col_ColumnComment).Value = IIf(Len(tableColumn.Comment) > 0, "'" & tableColumn.Comment, "")
        End If
        
        '-- Move next record
        Row = Row + 1
    Next
    
    '-- keep 2 blank column rows
    If clearExistedData Then
        For Row = Row To Row + 2
            If Row > Table_Sheet_Row_First_Column Then
                Call InsertColumnRow(sh, Row)
                SetColumnEmpty sh, Row
            End If
        Next
    End If
    
    '-- delete left column rows
    If clearExistedData Then
        Row = Row - 1
        For Row = Row To 32667
            If IsColumnRow(sh, Row) Then
                sh.Rows(Row & ":" & Row).Select
                Selection.Delete Shift:=xlUp
                Row = Row - 1
            Else
                Exit For
            End If
        Next
    End If
    sh.Cells(1, 1).Select
End Sub

Private Sub InsertColumnRow(sh As Worksheet, Row As Integer)
    
    sh.Rows(Row & ":" & Row).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    RenderRangeMerge sh.Range(Table_Sheet_Col_ColumnComment & Row & ":" & Table_Sheet_Col_TableStatus & Row)

End Sub

Private Function IsColumnRow(sh As Worksheet, Row As Integer) As Boolean
    Dim ma
    Set ma = sh.Range(Table_Sheet_Col_ColumnComment & Row).MergeArea
    IsColumnRow = (ma.Address = "$" & Table_Sheet_Col_ColumnComment & "$" & Row _
                    & ":$" & Table_Sheet_Col_TableStatus & "$" & Row)
End Function

Private Function SetColumnEmpty(ByRef sh As Worksheet, _
                        ByVal Row As Integer) As Boolean
    sh.Range(Table_Sheet_Col_ColumnLabel & Row, GetPreviousColumnName(Table_Sheet_Col_ColumnDataType) & Row).Select
    Selection.Interior.ColorIndex = xlNone
    Selection.Font.Bold = False
    '-- set Column
    sh.Cells(Row, Table_Sheet_Col_ColumnLabel).Value = ""
    sh.Cells(Row, Table_Sheet_Col_ColumnComment).Value = ""
    sh.Cells(Row, Table_Sheet_Col_ColumnName).Value = ""
    sh.Cells(Row, Table_Sheet_Col_ColumnDataType).Value = ""
    sh.Cells(Row, Table_Sheet_Col_ColumnNullable).Value = ""
    sh.Cells(Row, Table_Sheet_Col_ColumnDefault).Value = ""

End Function

Public Function SetForeignKeyRefTableAndName(ByRef foreignKey As clsLogicalForeignKey, _
                                            ByVal strRefTableAndColumns As String)
    Dim RefTableName    As String
    Dim refColumns      As String
    Dim fkOption        As String
    
    strRefTableAndColumns = Trim(strRefTableAndColumns)
    
    RefTableName = GetStringBefore(strRefTableAndColumns, "(")
    refColumns = GetStringAfter(strRefTableAndColumns, "(")
    fkOption = GetStringAfter(refColumns, ")")
    refColumns = GetStringBefore(refColumns, ")")
    
    foreignKey.RefTableName = RefTableName
    foreignKey.RefTableColumns = refColumns
    foreignKey.fkOption = fkOption
End Function

Public Function GetSheetFromTableName(ByVal TableName As String) As Worksheet
    Dim sh      As Worksheet
    Dim index   As Integer
    
    TableName = LCase(Trim(TableName))
    For index = Sheet_First_Table To ThisWorkbook.Sheets.Count
        If LCase(ThisWorkbook.Sheets(index).Cells(Table_Sheet_Row_TableName, Table_Sheet_Col_TableName).Text) = TableName Then
            Set sh = ThisWorkbook.Sheets(index)
            GoTo Exit_Flag
        End If
    Next
    
Exit_Flag:
    Set GetSheetFromTableName = sh
End Function
