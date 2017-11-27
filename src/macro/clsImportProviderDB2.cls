VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsImportProviderDB2"
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

Const ProviderOleDb = "IBMDADB2.DB2COPY1"
Const ProviderOdbc = "{IBM DB2 ODBC DRIVER}"
            
Public Property Get IImportProvider_Providers() As String()
    Dim candidateProviders(1 To 2) As String
    candidateProviders(1) = ProviderOleDb
    candidateProviders(2) = ProviderOdbc
    
    IImportProvider_Providers = candidateProviders
End Property

Public Property Get IImportProvider_SupportSelectDatabase() As Boolean
    IImportProvider_SupportSelectDatabase = True
End Property

Public Function IImportProvider_GetDatabasesSql() As String
    IImportProvider_GetDatabasesSql = ""
End Function

Public Function IImportProvider_GetDatabases(ByVal Server As String, _
                    ByVal UserID As String, _
                    ByVal Password As String) As Collection
    Dim databases As Collection
    Set databases = New Collection
    Dim batchFileContent As String
    
    Dim tempPath As String
    tempPath = VBA.Environ$("Temp") & "\~DME"
    basFile.DeleteFolder tempPath
    basFile.MakeFolder tempPath
    
    Dim batchFileName As String
    batchFileName = tempPath & "\GetDB2DatabaseNames.bat"
    Dim resultFileName As String
    resultFileName = tempPath & "\GetDB2DatabaseNames.txt"
    
    batchFileContent = "set db2instance=""" & Server & """"
    batchFileContent = batchFileContent _
                    & vbCrLf & "db2 list db directory > " & resultFileName
    basFile.SaveToTextFile batchFileName _
            , batchFileContent
            
    basShell.RunCommand "db2cmd -c -w -i " & batchFileName
    
    '-- Get result
    Dim result As String
    'Application.Wait (Now + TimeValue("0:00:02"))
    result = basFile.ReadFromTextFileFso(resultFileName)
    
    Dim results() As String
    Dim textLine As String
    Dim databaseName As String
    
    results = Split(result, vbLf)
    Dim i As Integer
    For i = LBound(results) To UBound(results)
        textLine = results(i)
        If LCase(basString.TrimEx(basString.GetStringBefore(textLine, "="))) = LCase("Database name") Then
            databaseName = basString.TrimEx(basString.GetStringAfter(textLine, "="))
            databases.Add databaseName, databaseName
        End If
    Next
    Set IImportProvider_GetDatabases = databases
End Function

Public Function IImportProvider_GetTablesSql() As String
    Dim syntax As String
    
    syntax = "  SELECT TABNAME as Name" _
    & Line & "    FROM syscat.tables" _
    & Line & "   WHERE OWNERTYPE='U'" _
    & Line & "     AND TYPE='T'" _
    & Line & "     AND TABSCHEMA=CURRENT_SCHEMA" _
    & Line & "ORDER BY TABNAME;"
    
    IImportProvider_GetTablesSql = syntax
End Function

Public Function IImportProvider_GetOptions() As clsImportOptions
    If basImport.mImportOptions Is Nothing Then
        Set basImport.mImportOptions = New clsImportOptions
        basImport.mImportOptions.DataSource = ""
        basImport.mImportOptions.DataSourceTip = "Leave it as <empty> for localhost or server name or IP."
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
    Const ConnectionStringFormatOleDb = "Provider={0: provider}" _
            & ";User ID={2: user}" _
            & ";Password={3: password}" _
            & ";Data Source={4: database}" _
            & ";Location={1: server}"
    Const ConnectionStringFormatOdbc = "Driver={0: provider}" _
            & ";Uid={2: user}" _
            & ";Pwd={3: password}" _
            & ";Database={4: database}" _
            & ";Hostname={1: server}" _
            & ";Protocol=TCPIP" _
            & ";Port=50000"
    Const Trused_ConnectionOption = ";Extended Properties=""trusted_connection=yes"""
    
    If Len(Trim(Database)) = 0 Then
        'Err.Raise 1, , "Please specific a database name."
    End If
    
    Dim connectionString As String
    Provider = Trim(Provider)
    
    If InStr(1, Provider, " ODBC ") Then
        connectionString = FormatString("Driver={0: provider}", Provider)
        connectionString = connectionString & ";Persist Security Info=False"
        connectionString = connectionString & FormatString(";Database={0: database}", Trim(Database))
        connectionString = connectionString & FormatString(";Hostname={0: server}", Trim(Server))
        If Len(Trim(User)) = 0 Then
            connectionString = connectionString & FormatString(";Uid={0: user}", Trim(User))
            connectionString = connectionString & ";Extended Properties=""trusted_connection=yes"""
        Else
            connectionString = connectionString & FormatString(";Uid={0: user}", Trim(User))
            connectionString = connectionString & FormatString(";Pwd={0: password}", Trim(Password))
            connectionString = connectionString & ";Extended Properties=""trusted_connection=yes"""
        End If
        connectionString = connectionString & ";Protocol=TCPIP" _
                            & ";Port=50000"
    Else
        connectionString = FormatString("Provider={0: provider}", Provider)
        connectionString = connectionString & ";Persist Security Info=False"
        connectionString = connectionString & FormatString(";Data Source={0: database}", Trim(Database))
        connectionString = connectionString & FormatString(";Location={0: server}", Trim(Server))
        If Len(Trim(User)) = 0 Then
            connectionString = connectionString & FormatString(";User ID={0: user}", Trim(User))
            connectionString = connectionString & ";Extended Properties=""trusted_connection=yes"""
        Else
            connectionString = connectionString & FormatString(";User ID={0: user}", Trim(User))
            connectionString = connectionString & FormatString(";Password={0: password}", Trim(Password))
            connectionString = connectionString & ";Extended Properties=""trusted_connection=yes"""
        End If
    End If

    IImportProvider_CreateConnectionString = connectionString
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
    syntax = "WITH IndexColumns (indid, indexColumnID, ColName, remainder, TabName, IsUnique, isPK, IsCluster)" _
    & Line & "AS" _
    & Line & "(" _
    & Line & "SELECT ROW_NUMBER() OVER () AS indid" _
    & Line & "     , 0" _
    & Line & "     , CASE WHEN LOCATE ('+', COLNAMES) > 0 THEN" _
    & Line & "         SUBSTR (COLNAMES, 1, LOCATE ('+', COLNAMES) - 1)" _
    & Line & "       Else" _
    & Line & "         COLNAMES" _
    & Line & "       END ColName" _
    & Line & "     , CASE WHEN locate ('+', COLNAMES) > 0 THEN" _
    & Line & "         LTRIM (SUBSTR (COLNAMES, LOCATE ('+', COLNAMES) + 1))" _
    & Line & "       Else" _
    & Line & "         NULL" _
    & Line & "       END remainder" _
    & Line & "     , TabName"
    syntax = syntax _
    & Line & "     , CASE WHEN UniqueRule = 'U' OR UniqueRule = 'P' THEN" _
    & Line & "         1" _
    & Line & "       Else" _
    & Line & "         0" _
    & Line & "       END As IsUnique" _
    & Line & "     , CASE WHEN UniqueRule = 'P' THEN" _
    & Line & "         1" _
    & Line & "       Else" _
    & Line & "         0" _
    & Line & "       END As IsPK"
    syntax = syntax _
    & Line & "     , CASE WHEN INDEXTYPE = 'CLUS' THEN" _
    & Line & "         1" _
    & Line & "       Else" _
    & Line & "         0" _
    & Line & "       END as IsCluster" _
    & Line & "  FROM SYSCAT.indexes base" _
    & Line & " WHERE TABSCHEMA = current_schema" _
    & Line & "   AND INDSCHEMA = current_schema" _
    & Line & "   AND OWNERTYPE = 'U'" _
    & Line & "   AND TABNAME='{0: table name in upper case}'" _
    & Line & "Union ALL"
    syntax = syntax _
    & Line & "SELECT indid" _
    & Line & "     , indexColumnID + 1" _
    & Line & "     , CASE WHEN LOCATE ('+', remainder) > 0 then" _
    & Line & "         SUBSTR (remainder, 1, LOCATE ('+', remainder) - 1)" _
    & Line & "       Else" _
    & Line & "         remainder" _
    & Line & "       END ColName" _
    & Line & "     , CASE WHEN LOCATE ('+', remainder) > 0 THEN" _
    & Line & "         LTRIM (SUBSTR (remainder, LOCATE ('+', remainder) + 1))" _
    & Line & "       Else" _
    & Line & "         NULL" _
    & Line & "       END remainder" _
    & Line & "     , TabName" _
    & Line & "     , IsUnique" _
    & Line & "     , isPK" _
    & Line & "     , IsCluster" _
    & Line & "  FROM IndexColumns t0" _
    & Line & " WHERE t0.remainder Is Not Null" _
    & Line & ")" _
    & Line & "  SELECT *" _
    & Line & "    FROM IndexColumns IC" _
    & Line & "   WHERE ColName IS NOT NULL" _
    & Line & "ORDER BY indid, indexColumnID;"

    Dim sSQL                    As String
    sSQL = FormatString(syntax, UCase(objTable.TableName))
    
    Dim oRs                     As ADODB.Recordset
    Dim curIndexID              As Integer
    Dim objIndex                As clsLogicalIndex

    On Error GoTo Flag_Err

    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    curIndexID = 0

    Do While Not oRs.EOF
        If oRs("isPK") = 1 Then
            '-- Primary Key
            If Len(objTable.PrimaryKey.PKcolumns) = 0 Then
                objTable.PrimaryKey.PKcolumns = oRs("ColName") & ""
            Else
                objTable.PrimaryKey.PKcolumns = objTable.PrimaryKey.PKcolumns & ", " & oRs("ColName")
            End If

            objTable.PrimaryKey.IsClustered = True
        Else
            '-- Index
            If curIndexID <> oRs("indid") Then
                Set objIndex = New clsLogicalIndex
                objTable.Indexes.Add objIndex
                
                objIndex.IsClustered = (oRs("IsCluster") = 1)
                objIndex.IsUnique = (oRs("IsUnique") = 1)

                curIndexID = oRs("indid")
            End If

            If Len(objIndex.IKColumns) = 0 Then
                objIndex.IKColumns = oRs("ColName") & ""
            Else
                objIndex.IKColumns = objIndex.IKColumns & ", " & oRs("ColName")
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
    Set oRs = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub RenderForeignKey(conn As ADODB.Connection, objTable As clsLogicalTable)
    Dim syntax As String
    
    syntax = "WITH RefInfo ( RefID" _
    & Line & "             , RefColumnID" _
    & Line & "             , ColName" _
    & Line & "             , Remainder" _
    & Line & "             , RefColName" _
    & Line & "             , RefRemainder" _
    & Line & "             , TabName" _
    & Line & "             , RefTabName" _
    & Line & "             , OnUpdate" _
    & Line & "             , OnDelete)" _
    & Line & "AS" _
    & Line & "(" _
    & Line & "  SELECT ROW_NUMBER() OVER () AS RefID" _
    & Line & "       , 0" _
    & Line & "       , CASE WHEN LOCATE (' ', FK_COLNAMES) > 0 THEN" _
    & Line & "           SUBSTR (FK_COLNAMES, 1, LOCATE (' ', FK_COLNAMES) - 1)" _
    & Line & "         Else" _
    & Line & "           FK_COLNAMES" _
    & Line & "         END ColName" _
    & Line & "       , CASE WHEN LOCATE (' ', FK_COLNAMES) > 0 THEN" _
    & Line & "           LTRIM (SUBSTR (FK_COLNAMES, LOCATE (' ', FK_COLNAMES) + 1))" _
    & Line & "         Else" _
    & Line & "           NULL" _
    & Line & "         END Remainder"
    syntax = syntax _
    & Line & "       , CASE WHEN LOCATE (' ', PK_COLNAMES) > 0 THEN" _
    & Line & "           SUBSTR (PK_COLNAMES, 1, LOCATE (' ', PK_COLNAMES) - 1)" _
    & Line & "         Else" _
    & Line & "           PK_COLNAMES" _
    & Line & "         END RefColName" _
    & Line & "       , CASE WHEN LOCATE (' ', PK_COLNAMES) > 0 THEN" _
    & Line & "           LTRIM (SUBSTR (PK_COLNAMES, LOCATE (' ', PK_COLNAMES) + 1))" _
    & Line & "         Else" _
    & Line & "           NULL" _
    & Line & "         END RefRemainder" _
    & Line & "       , TabName" _
    & Line & "       , RefTabName" _
    & Line & "       , UpdateRule as OnUpdate" _
    & Line & "       , DeleteRule as OnDelete" _
    & Line & "    FROM SYSCAT.REFERENCES base" _
    & Line & "   WHERE TABSCHEMA = current_schema" _
    & Line & "     AND OWNERTYPE = 'U'" _
    & Line & "     AND TABNAME='{0: table name in upper case}'"
    syntax = syntax _
    & Line & "  Union ALL" _
    & Line & "  SELECT RefID" _
    & Line & "       , RefColumnID + 1" _
    & Line & "       , CASE WHEN LOCATE (' ', Remainder) > 0 THEN" _
    & Line & "           SUBSTR (Remainder, 1, LOCATE (' ', Remainder) - 1)" _
    & Line & "         Else" _
    & Line & "           Remainder" _
    & Line & "         END ColName" _
    & Line & "       , CASE WHEN LOCATE (' ', Remainder) > 0 THEN" _
    & Line & "           LTRIM (SUBSTR (Remainder, LOCATE (' ', Remainder) + 1))" _
    & Line & "         Else" _
    & Line & "           NULL" _
    & Line & "         END Remainder" _
    & Line & "       , CASE WHEN LOCATE (' ', RefRemainder) > 0 then" _
    & Line & "           SUBSTR (RefRemainder, 1, LOCATE (' ', RefRemainder) - 1)" _
    & Line & "         Else" _
    & Line & "           Remainder" _
    & Line & "         END RefColName" _
    & Line & "       , CASE WHEN LOCATE (' ', RefRemainder) > 0 THEN" _
    & Line & "           LTRIM (SUBSTR (RefRemainder, LOCATE (' ', RefRemainder) + 1))" _
    & Line & "         Else" _
    & Line & "           NULL" _
    & Line & "         END RefRemainder"
    syntax = syntax _
    & Line & "       , TabName" _
    & Line & "       , RefTabName" _
    & Line & "       , OnUpdate" _
    & Line & "       , OnDelete" _
    & Line & "   FROM RefInfo t0" _
    & Line & "  WHERE t0.Remainder IS NOT NULL" _
    & Line & ")" _
    & Line & "  SELECT *" _
    & Line & "    FROM RefInfo" _
    & Line & "   WHERE colName IS NOT NULL" _
    & Line & "     AND Remainder IS NOT NULL" _
    & Line & "ORDER BY RefID, RefColumnID;"
    
    Dim sSQL                    As String
    sSQL = FormatString(syntax, UCase(objTable.TableName))
    
    Dim oRs             As ADODB.Recordset
    Dim curFKID         As String
    Dim objForeignKey   As clsLogicalForeignKey
    
    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    curFKID = ""

    Do While Not oRs.EOF
        '-- For Foreign Key
        If curFKID <> oRs("RefID") Then
            Set objForeignKey = New clsLogicalForeignKey
            objTable.ForeignKeys.Add objForeignKey

            objForeignKey.RefTableName = oRs("RefTabName") & ""
            objForeignKey.OnDelete = GetForeignKeyAction(oRs("OnDelete"), "ON DELETE")
            objForeignKey.OnUpdate = GetForeignKeyAction(oRs("OnUpdate"), "ON UPDATE")
            
            curFKID = oRs("RefID")
        End If

        If Len(objForeignKey.FKcolumns) > 0 Then
            objForeignKey.FKcolumns = objForeignKey.FKcolumns & ", "
        End If
        objForeignKey.FKcolumns = objForeignKey.FKcolumns & oRs("ColName")
        
        If Len(objForeignKey.RefTableColumns) > 0 Then
            objForeignKey.RefTableColumns = objForeignKey.RefTableColumns & ", "
        End If
        objForeignKey.RefTableColumns = objForeignKey.RefTableColumns & oRs("RefColName")

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
    
    syntax = "     SELECT C.COLNO" _
    & Line & "        , C.COLNAME as column_name" _
    & Line & "        , CASE WHEN C.NULLS = 'Y' THEN 1" _
    & Line & "          Else 0" _
    & Line & "          END As is_nullable" _
    & Line & "        , C.TYPENAME as data_type" _
    & Line & "        , C.STRINGUNITSLENGTH as max_string_length" _
    & Line & "        , C.TYPESTRINGUNITS as string_unit" _
    & Line & "        , C.LENGTH as numeric_precision" _
    & Line & "        , C.SCALE as numeric_scale" _
    & Line & "        , C.DEFAULT as column_default" _
    & Line & "        , C.REMARKS as column_comment" _
    & Line & "        , CA.START as identity_start" _
    & Line & "        , CA.INCREMENT as identity_increment" _
    & Line & "        , T.REMARKS as table_comment" _
    & Line & "     FROM SYSCAT.Columns C"
    syntax = syntax _
    & Line & "     JOIN SYSCAT.Tables T" _
    & Line & "       ON T.TABSCHEMA = C.TABSCHEMA" _
    & Line & "      AND T.TABNAME = C.TABNAME" _
    & Line & "LEFT JOIN SYSCAT.COLIDENTATTRIBUTES CA" _
    & Line & "       ON CA.TABSCHEMA = C.TABSCHEMA" _
    & Line & "      AND CA.TABNAME = C.TABNAME" _
    & Line & "      AND CA.COLNAME = C.COLNAME" _
    & Line & "      AND C.IDENTITY = 'Y'" _
    & Line & "   WHERE C.TABSCHEMA = Current_Schema" _
    & Line & "     AND C.TABNAME = '{0: table name in upper case}'" _
    & Line & "ORDER BY C.TABNAME" _
    & Line & "       , C.COLNO"

    Dim sSQL                    As String
    sSQL = FormatString(syntax, UCase(objTable.TableName))
    
    Dim oRs             As ADODB.Recordset
    Dim objColumn       As clsLogicalColumn
    
    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    Do While Not oRs.EOF
        '-- set Column
        Set objColumn = New clsLogicalColumn
        objTable.Columns.Add objColumn
        
        objColumn.ColumnName = oRs("column_name") & ""
        objColumn.DataType = GetColumnDataType( _
                                            oRs("data_type"), _
                                            IIf(IsNull(oRs("max_string_length")), -1, oRs("max_string_length")), _
                                            oRs("string_unit") & "", _
                                            oRs("numeric_precision"), _
                                            oRs("numeric_scale"), _
                                            IIf(IsNull(oRs("identity_start")), 0, oRs("identity_start")), _
                                            IIf(IsNull(oRs("identity_increment")), 0, oRs("identity_increment")))
        objColumn.Nullable = oRs("is_nullable")
        objColumn.Default = oRs("column_default") & ""
        objColumn.DBDescription = oRs("column_comment") & ""
        objTable.Comment = oRs("table_comment") & ""
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
                        maxLength As Long, _
                        stringUnit As String, _
                        precision As Long, _
                        type_scale As Integer, _
                        start As Long, _
                        increment As Long) As String
    Dim DataType As String

    DataType = LCase(typeName)
    Select Case DataType
        Case "char", "varchar", "binary", "varbinary", "nchar", "nvarchar", "graphic", "vargraphic", "character"
            If maxLength <> precision Then
                DataType = DataType & "(" & maxLength & " " & stringUnit & ")"
            Else
                DataType = DataType & "(" & maxLength & ")"
            End If
        Case "numeric", "decimal"
            If precision > 0 And type_scale >= 0 Then
                DataType = DataType & "(" & precision & ", " & type_scale & ")"
            End If
        Case "clob", "dbclob"
            If maxLength >= 0 And maxLength <> 1048576 Then
                DataType = DataType & "(" & GetLengthInUnit(maxLength) & ")"
            End If
    End Select

    If increment <> 0 Then
        DataType = DataType & FormatString(" GENERATED ALWAYS AS IDENTITY (START WITH {0: start}, INCREMENT BY {1: increment})", _
                        start, increment)
    End If
    GetColumnDataType = DataType
End Function

Private Function GetLengthInUnit(maxLength As Long) As String
    If maxLength Mod 2 ^ 30 = 0 Then
        GetLengthInUnit = maxLength / 2 ^ 30 & "G"
    ElseIf maxLength Mod 2 ^ 20 = 0 Then
        GetLengthInUnit = maxLength / 2 ^ 20 & "M"
    ElseIf maxLength Mod 2 ^ 10 = 0 Then
        GetLengthInUnit = maxLength / 2 ^ 10 & "K"
    Else
        GetLengthInUnit = maxLength
    End If

End Function

Private Function GetForeignKeyAction(action As String, _
            keyType As String)
    '-- a = no action, r = restrict, c = cascade, n = set null, d = set default
    Dim foreignKeyAction As String
    
    Select Case LCase(action)
    Case "r"
        foreignKeyAction = keyType & " RESTRICT"
    Case "c"
        foreignKeyAction = keyType & " CASCADE"
    Case "n"
        foreignKeyAction = keyType & " SET NULL"
    End Select
    
    GetForeignKeyAction = foreignKeyAction
End Function
