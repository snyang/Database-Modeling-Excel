VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsImportProviderOracle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2012, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
'-------------------------------------------
'-- Oracle Import Provider
'-------------------------------------------
Option Explicit
Implements IImportProvider

Public Property Get IImportProvider_Providers() As String()
    Dim candidateProviders(1 To 2) As String
    candidateProviders(1) = "MSDAORA.1"
    candidateProviders(2) = "OraOLEDB.Oracle.1"
    
    IImportProvider_Providers = candidateProviders
End Property

Public Property Get IImportProvider_SupportSelectDatabase() As Boolean
    IImportProvider_SupportSelectDatabase = False
End Property

Public Function IImportProvider_GetDatabasesSql() As String
    IImportProvider_GetDatabasesSql = ""
End Function

Public Function IImportProvider_GetDatabases(ByVal Server As String, _
                    ByVal UserID As String, _
                    ByVal Password As String) As Collection
    Set IImportProvider_GetDatabases = Nothing
End Function

Public Function IImportProvider_GetTablesSql() As String
    Dim syntax As String
    
    syntax = "  SELECT table_name AS name" _
    & Line & "    FROM User_Tables T" _
    & Line & "ORDER BY T.TABLE_NAME"
    
    IImportProvider_GetTablesSql = syntax
End Function

Public Function IImportProvider_GetOptions() As clsImportOptions
    If basImport.mImportOptions Is Nothing Then
        Set basImport.mImportOptions = New clsImportOptions
        basImport.mImportOptions.DataSource = "localhost/SID"
        basImport.mImportOptions.DataSourceTip = "e.g. localhost:1521/SID"
    End If
    Set IImportProvider_GetOptions = basImport.mImportOptions
End Function

Public Function IImportProvider_CreateConnectionString(ByVal Provider As String, _
                    ByVal Server As String, _
                    ByVal User As String, _
                    ByVal Password As String, _
                    ByVal Database As String) As String
                    
    IImportProvider_CreateConnectionString = "Provider=" & Trim(Provider) _
            & ";Data Source=" & Trim(Server) _
            & ";User ID=" & Trim(User) _
            & ";Password=" & Password _
            & ";Persist Security Info=True"
            
End Function

Public Function IImportProvider_GetLogicalTable(conn As ADODB.Connection, TableName As String) As clsLogicalTable
    Dim objTable As clsLogicalTable
    Set objTable = New clsLogicalTable
    
    objTable.TableName = TableName
    Set objTable.PrimaryKey = New clsLogicalPrimaryKey
    Set objTable.Indexes = New Collection
    Set objTable.ForeignKeys = New Collection
    Set objTable.Columns = New Collection
    
    RenderPKAndIndex conn, objTable
    RenderForeignKey conn, objTable
    RenderColumn conn, objTable
    
    '-- Return
    Set IImportProvider_GetLogicalTable = objTable
End Function

Public Sub RenderPKAndIndex(conn As ADODB.Connection, objTable As clsLogicalTable)
    Dim syntax As String
    
    syntax = "   SELECT TI.table_name" _
    & Line & "        , TI.index_name" _
    & Line & "        , TI.uniqueness" _
    & Line & "        , TIC.COLUMN_NAME" _
    & Line & "        , TIC.COLUMN_POSITION" _
    & Line & "        , TC.CONSTRAINT_TYPE" _
    & Line & "     FROM User_Indexes TI" _
    & Line & "     JOIN user_ind_columns TIC" _
    & Line & "       ON TIC.Table_Name = TI.table_name" _
    & Line & "      AND TIC.Index_Name = TI.index_name" _
    & Line & "LEFT JOIN user_constraints TC" _
    & Line & "       ON TI.table_name = TC.Table_Name" _
    & Line & "      AND TI.index_name = TC.Constraint_Name" _
    & Line & "    WHERE upper(TI.TABLE_NAME) = upper({0:table name})" _
    & Line & " ORDER BY TI.TABLE_NAME" _
    & Line & "        , TI.INDEX_NAME" _
    & Line & "        , TIC.Column_Position"

    Dim sSQL                    As String
    sSQL = FormatString(syntax, SQL_ToSQL(objTable.TableName))
    
    Dim oRs                     As ADODB.Recordset
    Dim curIndexName            As String
    Dim objIndex                As clsLogicalIndex

    On Error GoTo Flag_Err

    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    curIndexName = ""

    Do While Not oRs.EOF
        If oRs("CONSTRAINT_TYPE") = "P" Then
            '-- Primary Key
            If Len(objTable.PrimaryKey.PKcolumns) = 0 Then
                objTable.PrimaryKey.PKcolumns = oRs("COLUMN_NAME") & ""
            Else
                objTable.PrimaryKey.PKcolumns = objTable.PrimaryKey.PKcolumns & ", " & oRs("COLUMN_NAME")
            End If
            objTable.PrimaryKey.IsClustered = True
        Else
            '-- Index
            If curIndexName <> (oRs("INDEX_NAME") & "") Then
                Set objIndex = New clsLogicalIndex
                objTable.Indexes.Add objIndex
                
                objIndex.IsClustered = False
                objIndex.IsUnique = (oRs("uniqueness") = "UNIQUE")

                curIndexName = oRs("INDEX_NAME") & ""
            End If

            If Len(objIndex.IKColumns) = 0 Then
                objIndex.IKColumns = oRs("COLUMN_NAME") & ""
            Else
                objIndex.IKColumns = objIndex.IKColumns & ", " & oRs("COLUMN_NAME")
            End If
        End If

        '-- Move next record
        oRs.MoveNext
    Loop

    '-- Close record set
    oRs.Close
    Set oRs = Nothing
    Exit Sub
Flag_Err:
    If Not oRs Is Nothing Then oRs.Close
    Set oRs = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub RenderForeignKey(conn As ADODB.Connection, objTable As clsLogicalTable)
    Dim syntax As String
    
    syntax = "SELECT UC.TABLE_NAME" _
    & Line & "     , UC.CONSTRAINT_NAME" _
    & Line & "     , UC.delete_rule" _
    & Line & "     , UCC.column_name" _
    & Line & "     , UCC.POSITION" _
    & Line & "     , UCCR.table_name AS REFERENCED_TABLE_NAME" _
    & Line & "     , UCCR.column_name AS REFERENCED_COLUMN_NAME" _
    & Line & "  FROM user_constraints UC" _
    & Line & "  JOIN user_cons_columns UCC" _
    & Line & "    ON UC.TABLE_NAME        = UCC.TABLE_NAME" _
    & Line & "   AND UC.CONSTRAINT_NAME   = UCC.CONSTRAINT_NAME" _
    & Line & "  JOIN user_cons_columns UCCR" _
    & Line & "    ON UC.r_constraint_name   = UCCR.constraint_name" _
    & Line & "   AND UCC.position = UCCR.position" _
    & Line & " WHERE UC.TABLE_NAME = {0:table name}" _
    & Line & "   AND UC.constraint_type = 'R'" _
    & Line & " ORDER BY UC.TABLE_NAME" _
    & Line & "     , UC.CONSTRAINT_NAME" _
    & Line & "     , UCC.POSITION"

    Dim sSQL                    As String
    sSQL = FormatString(syntax, SQL_ToSQL(objTable.TableName))
    
    Dim oRs             As ADODB.Recordset
    Dim curFKName       As String
    Dim objForeignKey   As clsLogicalForeignKey
    
    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    curFKName = ""

    Do While Not oRs.EOF
        '-- For Foreign Key
        If curFKName <> (oRs("CONSTRAINT_NAME") & "") Then
            Set objForeignKey = New clsLogicalForeignKey
            objTable.ForeignKeys.Add objForeignKey

            objForeignKey.RefTableName = oRs("REFERENCED_TABLE_NAME") & ""
            If oRs("DELETE_RULE") <> "NO ACTION" Then
                objForeignKey.OnDelete = "ON DELETE " & oRs("DELETE_RULE")
            Else
                objForeignKey.OnDelete = ""
            End If
          
            objForeignKey.OnUpdate = ""
            
            curFKName = oRs("CONSTRAINT_NAME") & ""
        End If

        If Len(objForeignKey.FKcolumns) > 0 Then
            objForeignKey.FKcolumns = objForeignKey.FKcolumns & ", "
        End If
        objForeignKey.FKcolumns = objForeignKey.FKcolumns & oRs("COLUMN_NAME")
        
        If Len(objForeignKey.RefTableColumns) > 0 Then
            objForeignKey.RefTableColumns = objForeignKey.RefTableColumns & ", "
        End If
        objForeignKey.RefTableColumns = objForeignKey.RefTableColumns & oRs("REFERENCED_COLUMN_NAME")

        '-- Move next record
        oRs.MoveNext
    Loop

    '-- Close record set
    oRs.Close
    Set oRs = Nothing

    Exit Sub
Flag_Err:
    If Not oRs Is Nothing Then oRs.Close
    Set oRs = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub RenderColumn(conn As ADODB.Connection, objTable As clsLogicalTable)
    Dim syntax As String
    
    syntax = "    SELECT C.TABLE_NAME" _
    & Line & "         , C.COLUMN_NAME" _
    & Line & "         , C.COLUMN_ID" _
    & Line & "         , C.DATA_TYPE" _
    & Line & "         , C.NULLABLE" _
    & Line & "         , C.DATA_LENGTH" _
    & Line & "         , C.CHAR_LENGTH" _
    & Line & "         , C.DATA_PRECISION" _
    & Line & "         , C.DATA_SCALE" _
    & Line & "         , C.DATA_DEFAULT" _
    & Line & "         , TCC.COMMENTS as ColumnComments" _
    & Line & "         , TC.COMMENTS as TableComments" _
    & Line & "      FROM User_Tab_Columns C" _
    & Line & " LEFT JOIN User_Tab_Comments TC" _
    & Line & "        ON TC.TABLE_NAME = C.TABLE_NAME" _
    & Line & " LEFT JOIN User_Col_Comments TCC" _
    & Line & "        ON TCC.TABLE_NAME = C.TABLE_NAME" _
    & Line & "       AND TCC.COLUMN_NAME = C.COLUMN_NAME" _
    & Line & "     WHERE C.TABLE_NAME = {0:table name}" _
    & Line & "  ORDER BY C.TABLE_NAME" _
    & Line & "         , C.COLUMN_ID"

    Dim sSQL                    As String
    sSQL = FormatString(syntax, SQL_ToSQL(objTable.TableName))
    
    Dim oRs             As ADODB.Recordset
    Dim objColumn       As clsLogicalColumn
    
    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    Do While Not oRs.EOF
        '-- set Column
        Set objColumn = New clsLogicalColumn
        objTable.Columns.Add objColumn
        
        objColumn.ColumnName = oRs("COLUMN_NAME") & ""
        objColumn.DataType = GetColumnDataType( _
                                            oRs("DATA_TYPE"), _
                                            oRs("DATA_LENGTH") & "", _
                                            oRs("CHAR_LENGTH") & "", _
                                            oRs("DATA_PRECISION") & "", _
                                            oRs("DATA_SCALE") & "")
        objColumn.Nullable = (oRs("NULLABLE") = "Y")
        If Not IsNull(oRs("DATA_DEFAULT")) Then
            objColumn.Default = oRs("DATA_DEFAULT")
        Else
            objColumn.Default = ""
        End If
        
        If Not IsNull(oRs("ColumnComments")) Then
            objColumn.DBDescription = oRs("ColumnComments")
        Else
            objColumn.DBDescription = ""
        End If
        
        If Not IsNull(oRs("TableComments")) Then
            objTable.Comment = oRs("TableComments")
        Else
            objTable.Comment = ""
        End If
        
        '-- Move next record
        oRs.MoveNext
    Loop

    '-- Close record set
    oRs.Close
    Set oRs = Nothing

    Exit Sub
Flag_Err:
    If Not oRs Is Nothing Then oRs.Close
    Set oRs = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Function GetColumnDataType(typeName As String, _
                        maxLength As String, _
                        charLength As String, _
                        precision As String, _
                        type_scale As String) As String
    Dim DataType As String

    DataType = LCase(typeName)
    Select Case DataType
        Case "varchar2", "nvarchar2", "char", "nchar"
            DataType = DataType & "(" & charLength & ")"
        Case "raw"
            If CInt(maxLength) > 0 Then
                DataType = DataType & "(" & maxLength & ")"
            End If
        Case "float"
            If CInt(precision) >= 126 Then
                DataType = DataType
            ElseIf precision = "" _
                And (type_scale = "0" Or type_scale = "") Then
                DataType = DataType
            ElseIf type_scale = "0" Or type_scale = "" Then
                DataType = DataType & "(" & precision & ")"
            Else
                DataType = DataType & "(" & precision & ", " & type_scale & ")"
            End If
        Case "numeric", "number"
            If precision = "" _
                And (type_scale = "0" Or type_scale = "") Then
                DataType = DataType
            ElseIf type_scale = "0" Or type_scale = "" Then
                DataType = DataType & "(" & precision & ")"
            Else
                DataType = DataType & "(" & precision & ", " & type_scale & ")"
            End If
    End Select

    GetColumnDataType = DataType
End Function



