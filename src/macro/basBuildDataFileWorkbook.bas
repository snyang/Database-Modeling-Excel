Attribute VB_Name = "basBuildDataFileWorkbook"
Option Explicit

Private Const Type_Table = "T"
Private Const Type_PK = "P"
Private Const Type_Index = "I"
Private Const Type_FK = "F"
Private Const Type_Field = "C"

Public Sub ExportToWorkbook()
    Dim objWorkbook As Workbook
    Set objWorkbook = Application.Workbooks.Add
    Call EmptyWorkbook(objWorkbook)
    
    Dim objTables As Collection
    Set objTables = GetAllLogicalTables(True)

    Dim objTable        As clsLogicalTable
    Dim objIndex        As clsLogicalIndex
    Dim objForeignKey   As clsLogicalForeignKey
    Dim objSheet        As Worksheet
    Dim objColumn       As clsLogicalColumn
    
    If objWorkbook.Sheets.Count = 1 Then
        Set objSheet = objWorkbook.Sheets(1)
    Else
        Set objSheet = objWorkbook.Sheets.Add(, objWorkbook.Sheets(objWorkbook.Sheets.Count))
    End If
    objSheet.Name = The_Excel_Type
    
    Dim iRow            As Integer
    iRow = 1
    For Each objTable In objTables
        
        '-- table
        objSheet.Cells(iRow, "A").Value = Type_Table
        objSheet.Cells(iRow, "B").Value = objTable.TableName
        objSheet.Cells(iRow, "C").Value = objTable.Comment
        
        '-- PK
        If Not objTable.PrimaryKey Is Nothing Then
            iRow = iRow + 1
            objSheet.Cells(iRow, "A").Value = Type_PK
            objSheet.Cells(iRow, "B").Value = objTable.PrimaryKey.PKcolumns
            objSheet.Cells(iRow, "C").Value = objTable.PrimaryKey.IsClustered
        End If
        
        '-- Index
        For Each objIndex In objTable.Indexes
            iRow = iRow + 1
            objSheet.Cells(iRow, "A").Value = Type_Index
            objSheet.Cells(iRow, "B").Value = Replace(Replace(objIndex.IKColumns, "(", ""), ")", "")
            objSheet.Cells(iRow, "C").Value = objIndex.IsClustered
            objSheet.Cells(iRow, "D").Value = objIndex.IsUnique
        Next
        
        '-- FK
        For Each objForeignKey In objTable.ForeignKeys
            iRow = iRow + 1
            objSheet.Cells(iRow, "A").Value = Type_FK
            objSheet.Cells(iRow, "B").Value = objForeignKey.FKcolumns
            objSheet.Cells(iRow, "C").Value = objForeignKey.RefTableName
            objSheet.Cells(iRow, "D").Value = objForeignKey.RefTableColumns
            objSheet.Cells(iRow, "E").Value = objForeignKey.fkOption
        Next
        
        '-- Columns
        For Each objColumn In objTable.Columns
            iRow = iRow + 1
            objSheet.Cells(iRow, "A").Value = Type_Field
            objSheet.Cells(iRow, "B").Value = objColumn.ColumnLabel
            objSheet.Cells(iRow, "C").Value = objColumn.ColumnName
            objSheet.Cells(iRow, "D").Value = objColumn.DataType
            objSheet.Cells(iRow, "E").Value = objColumn.Nullable
            objSheet.Cells(iRow, "F").Value = "'" & objColumn.Default
            objSheet.Cells(iRow, "G").Value = objColumn.Comment
        Next
        iRow = iRow + 1
    Next
    objSheet.Cells.EntireColumn.AutoFit
End Sub

Private Sub EmptyWorkbook(objWorkbook As Workbook)
    While objWorkbook.Sheets.Count > 1
        objWorkbook.Sheets(2).Delete
    Wend
End Sub

Public Function ImportHasTable(sheetData As Worksheet, _
                    ByRef Row As Integer)
    ImportHasTable = Len(sheetData.Cells(Row, "A")) > 0
End Function

'---------------------------------------------
'-  Write Table information to worksheet
'---------------------------------------------
Public Sub FillTableDefinitionFromDataSheet(ByVal sheetData As Worksheet, _
                        ByVal sheetTable As Worksheet, _
                        ByRef Row As Integer)
    Dim objTable As clsLogicalTable
    Set objTable = New clsLogicalTable
    
    '-- Set Table Name
    If sheetData.Cells(Row, "A") = Type_Table Then
        objTable.TableName = sheetData.Cells(Row, "B").Value
        objTable.Comment = sheetData.Cells(Row, "C").Value
    Else
        Exit Sub
    End If
    
    Dim objForeignKey As clsLogicalForeignKey
    Dim objIndex As clsLogicalIndex
    Dim objColumn As clsLogicalColumn
    
    Do While True
        Row = Row + 1
        If sheetData.Cells(Row, "A") = Type_Table Then
            Exit Do
            
        ElseIf sheetData.Cells(Row, "A") = Type_PK Then
            '-- PK
            Set objTable.PrimaryKey = New clsLogicalPrimaryKey
            objTable.PrimaryKey.PKcolumns = sheetData.Cells(Row, "B").Value
            objTable.PrimaryKey.IsClustered = sheetData.Cells(Row, "C").Value
            
        ElseIf sheetData.Cells(Row, "A") = Type_FK Then
            '-- FK
            Set objForeignKey = New clsLogicalForeignKey
            objForeignKey.FKcolumns = sheetData.Cells(Row, "B").Value
            objForeignKey.RefTableName = sheetData.Cells(Row, "C").Value
            objForeignKey.RefTableColumns = sheetData.Cells(Row, "D").Value
            objForeignKey.OnDelete = sheetData.Cells(Row, "E").Value
            objForeignKey.OnUpdate = sheetData.Cells(Row, "F").Value
            objTable.ForeignKeys.Add objForeignKey
            
        ElseIf sheetData.Cells(Row, "A") = Type_Index Then
            '-- Index
            Set objIndex = New clsLogicalIndex
            objIndex.IKColumns = sheetData.Cells(Row, "B").Value
            objIndex.IsClustered = sheetData.Cells(Row, "C").Value
            objIndex.IsUnique = sheetData.Cells(Row, "D").Value
            objTable.Indexes.Add objIndex
            
        ElseIf sheetData.Cells(Row, "A") = Type_Field Then
            '-- Column
            Set objColumn = New clsLogicalColumn
            objColumn.ColumnLabel = sheetData.Cells(Row, "B").Value
            objColumn.ColumnName = sheetData.Cells(Row, "C").Value
            objColumn.DataType = sheetData.Cells(Row, "D").Value
            objColumn.Nullable = sheetData.Cells(Row, "E").Value
            objColumn.Default = StripBrackets(sheetData.Cells(Row, "F").Value)
            objColumn.Comment = sheetData.Cells(Row, "G").Value
            objTable.Columns.Add objColumn
            
        Else
            Exit Do
        End If
    Loop
 
    sheetTable.Activate
    basTableSheet.SetTableInfoToWorksheet sheetTable, objTable, True
End Sub
