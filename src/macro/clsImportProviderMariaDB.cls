VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsImportProviderMariaDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2014, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit
Implements IImportProvider

Public Property Get IImportProvider_Providers() As String()
    Const ProviderOdbc = "{MariaDB ODBC 1.0 Driver}"
    
    Dim candidateProviders(1 To 1) As String
    candidateProviders(1) = ProviderOdbc
    
    IImportProvider_Providers = candidateProviders
End Property

Public Property Get IImportProvider_SupportSelectDatabase() As Boolean
    IImportProvider_SupportSelectDatabase = True
End Property

Public Function IImportProvider_GetDatabasesSql() As String
    
    Dim syntax As String
    
    syntax = "  SELECT SCHEMA_NAME AS name" _
    & Line & "    FROM information_schema.SCHEMATA" _
    & Line & "ORDER BY SCHEMA_NAME;"

    IImportProvider_GetDatabasesSql = syntax
    
End Function

Public Function IImportProvider_GetDatabases(ByVal Server As String, _
                    ByVal UserID As String, _
                    ByVal Password As String) As Collection
    Set IImportProvider_GetDatabases = Nothing
End Function

Public Function IImportProvider_GetTablesSql() As String
    Dim syntax As String
    
    syntax = "  SELECT TABLE_NAME AS name" _
    & Line & "    FROM information_schema.TABLES" _
    & Line & "   WHERE TABLE_SCHEMA = DATABASE()" _
    & Line & "ORDER BY TABLE_NAME;"
    
    IImportProvider_GetTablesSql = syntax
End Function

Public Function IImportProvider_GetOptions() As clsImportOptions
    If basImport.mImportOptions Is Nothing Then
        Set basImport.mImportOptions = New clsImportOptions
        basImport.mImportOptions.DataSource = "localhost"
        basImport.mImportOptions.DataSourceTip = "Server name or IP."
        basImport.mImportOptions.UserName = ""
        basImport.mImportOptions.Database = ""
    End If
    Set IImportProvider_GetOptions = basImport.mImportOptions
End Function

Public Function IImportProvider_CreateConnectionString(ByVal Provider As String, _
                    ByVal Server As String, _
                    ByVal User As String, _
                    ByVal Password As String, _
                    ByVal Database As String) As String
    Const ConnectionStringFormatOdbc = "Driver={0: provider}" _
            & ";UID={2: user}" _
            & ";PWD={3: password}" _
            & ";DB={4: database}" _
            & ";SERVER={1: server}" _
            & ";PORT=3306"
    
    IImportProvider_CreateConnectionString = FormatString(ConnectionStringFormatOdbc, _
            Trim(Provider), _
            Trim(Server), _
            Trim(User), _
            Trim(Password), _
            Trim(Database))
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
    
    syntax = "   SELECT S.TABLE_NAME" _
    & Line & "        , S.INDEX_NAME" _
    & Line & "        , S.SEQ_IN_INDEX" _
    & Line & "        , S.COLUMN_NAME" _
    & Line & "        , S.NON_UNIQUE" _
    & Line & "        , TC.CONSTRAINT_TYPE" _
    & Line & "     FROM information_schema.STATISTICS S" _
    & Line & "LEFT JOIN information_schema.TABLE_CONSTRAINTS TC" _
    & Line & "       ON TC.TABLE_SCHEMA = S.TABLE_SCHEMA" _
    & Line & "      AND TC.TABLE_NAME = S.TABLE_NAME" _
    & Line & "      AND TC.CONSTRAINT_NAME = S.INDEX_NAME" _
    & Line & "    WHERE S.TABLE_SCHEMA = DATABASE()" _
    & Line & "      AND S.TABLE_NAME = {0:table name}" _
    & Line & "      AND (TC.CONSTRAINT_TYPE IS NULL OR TC.CONSTRAINT_TYPE IN ('PRIMARY KEY', 'UNIQUE'))" _
    & Line & " ORDER BY S.TABLE_NAME" _
    & Line & "        , S.INDEX_NAME" _
    & Line & "        , S.SEQ_IN_INDEX;"

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
        If oRs("CONSTRAINT_TYPE") & "" = "PRIMARY KEY" Then
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
                objIndex.IsUnique = (oRs("NON_UNIQUE") = 0)

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
    If Not oRs Is Nothing And oRs.State <> 0 Then oRs.Close
    Set oRs = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub RenderForeignKey(conn As ADODB.Connection, objTable As clsLogicalTable)
    Dim syntax As String
    
    syntax = "SELECT R.TABLE_NAME" _
    & Line & "     , R.CONSTRAINT_NAME" _
    & Line & "     , R.UPDATE_RULE" _
    & Line & "     , R.DELETE_RULE" _
    & Line & "     , R.REFERENCED_TABLE_NAME" _
    & Line & "     , K.COLUMN_NAME" _
    & Line & "     , K.ORDINAL_POSITION" _
    & Line & "     , K.POSITION_IN_UNIQUE_CONSTRAINT" _
    & Line & "     , K.REFERENCED_COLUMN_NAME" _
    & Line & "  FROM information_schema.REFERENTIAL_CONSTRAINTS R" _
    & Line & "  JOIN information_schema.KEY_COLUMN_USAGE K" _
    & Line & "    ON R.CONSTRAINT_SCHEMA = K.CONSTRAINT_SCHEMA" _
    & Line & "   AND R.TABLE_NAME        = K.TABLE_NAME" _
    & Line & "   AND R.CONSTRAINT_NAME   = K.CONSTRAINT_NAME" _
    & Line & " WHERE R.CONSTRAINT_SCHEMA = DATABASE()" _
    & Line & "   AND R.TABLE_NAME = {0:table name}" _
    & Line & " ORDER BY R.TABLE_NAME" _
    & Line & "     , R.CONSTRAINT_NAME" _
    & Line & "     , K.ORDINAL_POSITION;"

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

            objForeignKey.RefTableName = oRs("REFERENCED_TABLE_NAME")
            If oRs("DELETE_RULE") <> "RESTRICT" Then
                objForeignKey.OnDelete = "ON DELETE " & oRs("DELETE_RULE")
            Else
                objForeignKey.OnDelete = ""
            End If
            If oRs("UPDATE_RULE") <> "RESTRICT" Then
                objForeignKey.OnUpdate = "ON UPDATE " & oRs("UPDATE_RULE")
            Else
                objForeignKey.OnUpdate = ""
            End If
            
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
    
    syntax = "  SELECT T.TABLE_NAME" _
    & Line & "       , C.COLUMN_NAME" _
    & Line & "       , C.ORDINAL_POSITION" _
    & Line & "       , C.COLUMN_TYPE" _
    & Line & "       , C.COLUMN_DEFAULT" _
    & Line & "       , C.EXTRA" _
    & Line & "       , C.IS_NULLABLE" _
    & Line & "       , C.COLUMN_COMMENT" _
    & Line & "       , C.DATA_TYPE" _
    & Line & "       , C.CHARACTER_MAXIMUM_LENGTH" _
    & Line & "       , C.CHARACTER_SET_NAME" _
    & Line & "       , C.NUMERIC_PRECISION" _
    & Line & "       , C.NUMERIC_SCALE" _
    & Line & "       , T.TABLE_COMMENT" _
    & Line & "    FROM information_schema.TABLES T" _
    & Line & "    JOIN information_schema.COLUMNS C" _
    & Line & "      ON C.TABLE_SCHEMA = T.TABLE_SCHEMA" _
    & Line & "     AND C.TABLE_NAME = T.TABLE_NAME" _
    & Line & "   WHERE T.TABLE_SCHEMA = DATABASE()" _
    & Line & "     AND T.TABLE_TYPE = 'BASE TABLE'" _
    & Line & "     AND T.TABLE_NAME = {0:table name}" _
    & Line & "ORDER BY C.TABLE_NAME" _
    & Line & "       , C.ORDINAL_POSITION;"

    Dim sSQL                    As String
    sSQL = FormatString(syntax, SQL_ToSQL(objTable.TableName))
    
    Dim oRs             As ADODB.Recordset
    Dim objColumn       As clsLogicalColumn
    Dim defaultValue    As Variant
    
    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    Do While Not oRs.EOF
        '-- set Column
        Set objColumn = New clsLogicalColumn
        objTable.Columns.Add objColumn
        
        objColumn.ColumnName = oRs("COLUMN_NAME")
        objColumn.DataType = GetColumnDataType( _
                                            oRs("COLUMN_TYPE"), _
                                            oRs("CHARACTER_SET_NAME") & "", _
                                            oRs("EXTRA") & "")
        objColumn.Nullable = (oRs("IS_NULLABLE") = "YES")
        defaultValue = oRs("COLUMN_DEFAULT")
        objColumn.Default = GetDefaultValue(oRs("COLUMN_TYPE"), defaultValue)
        objColumn.DBDescription = oRs("COLUMN_COMMENT") & ""
        objTable.Comment = oRs("TABLE_COMMENT") & ""
        
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

Private Function GetColumnDataType(columnType As String, _
                        characterSetName As String, _
                        extra As String) As String
    Dim DataType As String

    DataType = columnType
'    If basString.StartWith(DataType, "char") _
'        Or basString.StartWith(DataType, "varchar") Then
'        If characterSetName = "utf8" Then
'            DataType = "n" & DataType
'        End If
'    Else
    If basString.StartWith(DataType, "year(4)") Then
        DataType = "year" & Mid(DataType, Len("year(4)") + 1)
        
    ElseIf basString.StartWith(DataType, "bigint(20)") Then
        DataType = "bigint" & Mid(DataType, Len("bigint(20)") + 1)
    
    ElseIf basString.StartWith(DataType, "bigint(19) unsigned") Then
        DataType = "bigint unsigned" & Mid(DataType, Len("bigint(20) unsigned") + 1)
    
    ElseIf basString.StartWith(DataType, "int(11)") Then
        DataType = "int" & Mid(DataType, Len("int(11)") + 1)
    
    ElseIf basString.StartWith(DataType, "int(10) unsigned") Then
        DataType = "int unsigned" & Mid(DataType, Len("int(10) unsigned") + 1)
        
    ElseIf basString.StartWith(DataType, "mediumint(9)") Then
        DataType = "mediumint" & Mid(DataType, Len("mediumint(9)") + 1)
    
    ElseIf basString.StartWith(DataType, "mediumint(8) unsigned") Then
        DataType = "mediumint unsigned" & Mid(DataType, Len("mediumint(8) unsigned") + 1)
    
    ElseIf basString.StartWith(DataType, "smallint(6)") Then
        DataType = "smallint" & Mid(DataType, Len("smallint(6)") + 1)
    
    ElseIf basString.StartWith(DataType, "smallint(5) unsigned") Then
        DataType = "smallint unsigned" & Mid(DataType, Len("smallint(5) unsigned") + 1)
    
    ElseIf basString.StartWith(DataType, "tinyint(4)") Then
        DataType = "tinyint" & Mid(DataType, Len("tinyint(4)") + 1)
    
    ElseIf basString.StartWith(DataType, "tinyint(3) unsigned") Then
        DataType = "tinyint unsigned" & Mid(DataType, Len("tinyint(3) unsigned") + 1)
    
    ElseIf basString.StartWith(DataType, "bit(1)") Then
        DataType = "bit" & Mid(DataType, Len("bit(1)") + 1)
    End If
        
    If (Len(extra) > 0) Then
        DataType = DataType & " " & extra
    End If
    GetColumnDataType = DataType
End Function

Private Function GetDefaultValue(DataType As String, _
                            defaultValue As Variant) As String
    GetDefaultValue = ""
    If Not IsNull(defaultValue) Then
        If basString.StartWith(DataType, "char") _
            Or basString.StartWith(DataType, "varchar") _
            Or basString.StartWith(DataType, "nchar") _
            Or basString.StartWith(DataType, "nvarchar") _
            Or basString.StartWith(DataType, "enum") _
            Then
            
            GetDefaultValue = "'" & defaultValue & "'"
        Else
            GetDefaultValue = defaultValue
        End If
    End If
End Function

