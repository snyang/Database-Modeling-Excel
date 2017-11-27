VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsImportProviderPostgreSQL"
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

Const ProviderOleDb = "PostgreSQL OLE DB Provider"
Const ConnectionStringFormatOleDb = "Provider={0: provider}" _
            & ";User ID={2: user}" _
            & ";password={3: password}" _
            & ";location={4: database}" _
            & ";Data Source={1: server}"
Const ProviderOdbc = "{PostgreSQL ODBC Driver(UNICODE)}"
Const ConnectionStringFormatOdbc = "Driver={0: provider}" _
            & ";Uid={2: user}" _
            & ";Pwd={3: password}" _
            & ";Database={4: database}" _
            & ";Server={1: server}" _
            & ";Port=5432"

Public Property Get IImportProvider_Providers() As String()
    Dim candidateProviders(1 To 2) As String
    candidateProviders(1) = ProviderOdbc
    candidateProviders(2) = ProviderOleDb
    
    IImportProvider_Providers = candidateProviders
End Property

Public Property Get IImportProvider_SupportSelectDatabase() As Boolean
    IImportProvider_SupportSelectDatabase = True
End Property

Public Function IImportProvider_GetDatabasesSql() As String
    Dim syntax As String
    
    syntax = "  SELECT datname AS name" _
    & Line & "    FROM pg_database" _
    & Line & "ORDER BY datname"
    
    IImportProvider_GetDatabasesSql = syntax
End Function

Public Function IImportProvider_GetDatabases(ByVal Server As String, _
                    ByVal UserID As String, _
                    ByVal Password As String) As Collection
    Set IImportProvider_GetDatabases = Nothing
End Function

Public Function IImportProvider_GetTablesSql() As String
    Dim syntax As String
    
    syntax = "  SELECT table_name AS name" _
    & Line & "    FROM information_schema.tables" _
    & Line & "   WHERE table_type = 'BASE TABLE'" _
    & Line & "     AND table_schema = 'public'" _
    & Line & "ORDER BY table_name"
    
    IImportProvider_GetTablesSql = syntax
End Function

Public Function IImportProvider_GetOptions() As clsImportOptions
    If basImport.mImportOptions Is Nothing Then
        Set basImport.mImportOptions = New clsImportOptions
        basImport.mImportOptions.DataSource = "localhost"
        basImport.mImportOptions.DataSourceTip = "e.g. localhost"
        basImport.mImportOptions.UserName = "postgres"
        basImport.mImportOptions.Database = "postgres"
    End If
    Set IImportProvider_GetOptions = basImport.mImportOptions
End Function

Public Function IImportProvider_CreateConnectionString(ByVal Provider As String, _
                    ByVal Server As String, _
                    ByVal User As String, _
                    ByVal Password As String, _
                    ByVal Database As String) As String
                    
    Dim connectionStringFormat As String
    Provider = Trim(Provider)
    If LCase(Provider) = LCase(ProviderOdbc) Then
        IImportProvider_CreateConnectionString = FormatString(ConnectionStringFormatOdbc, _
                        Provider, _
                        Trim(Server), _
                        Trim(User), _
                        Password, _
                        IIf(Len(Database) = 0, "postgres", Database))
    Else
        IImportProvider_CreateConnectionString = FormatString(ConnectionStringFormatOleDb, _
                        Provider, _
                        Trim(Server), _
                        Trim(User), _
                        Password, _
                        Database)
    End If
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
    syntax = "   SELECT c.oid" _
    & Line & "        , i.indexrelid as indid" _
    & Line & "        , i.indisunique as IsUnique" _
    & Line & "        , i.indisprimary as isPK" _
    & Line & "        , i.indisclustered as IsCluster" _
    & Line & "        , i.indkey" _
    & Line & "        , i.indattind" _
    & Line & "        , a.attnum" _
    & Line & "        , a.attname as ColName" _
    & Line & "     FROM pg_namespace AS nsp" _
    & Line & "     JOIN pg_class AS c" _
    & Line & "       ON c.relnamespace = nsp.oid" _
    & Line & "      AND c.relname = '{0:table name}'" _
    & Line & "     JOIN (" _
    & Line & "          SELECT indrelid" _
    & Line & "           , indexrelid" _
    & Line & "           , indisunique" _
    & Line & "           , indisprimary" _
    & Line & "           , indisclustered" _
    & Line & "           , indkey" _
    & Line & "           , generate_subscripts(indkey, 1) as indattind" _
    & Line & "           FROM pg_index" _
    & Line & "           ) i" _
    & Line & "       ON i.indrelid = c.oid" _
    & Line & "     JOIN pg_attribute a"
    syntax = syntax _
    & Line & "       ON a.attrelid = c.oid" _
    & Line & "      AND a.attnum > 0" _
    & Line & "      AND a.attnum = i.indkey[indattind]" _
    & Line & "    WHERE nsp.nspname = current_schema()" _
    & Line & " ORDER BY c.oid" _
    & Line & "        , i.indexrelid" _
    & Line & "        , i.indattind"

    Dim sSQL                    As String
    sSQL = FormatString(syntax, objTable.TableName)
    
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

            objTable.PrimaryKey.IsClustered = (oRs("IsCluster") = 1)
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
    
    syntax = "   SELECT c.oid" _
    & Line & "        , con.conname as id" _
    & Line & "        , con.conkey" _
    & Line & "        , con.confkey" _
    & Line & "        , con.confupdtype as OnUpdate" _
    & Line & "        , con.confdeltype as OnDelete" _
    & Line & "        , con.confmatchtype" _
    & Line & "        , a.attnum" _
    & Line & "        , a.attname as ftColumnName" _
    & Line & "        , cf.relname as rtName" _
    & Line & "        , af.attnum" _
    & Line & "        , af.attname as rtColumnName" _
    & Line & "     FROM pg_namespace AS nsp" _
    & Line & "     JOIN pg_class AS c" _
    & Line & "       ON c.relnamespace = nsp.oid" _
    & Line & "      AND c.relname = '{0:table name}'" _
    & Line & "     JOIN (" _
    & Line & "          SELECT conname" _
    & Line & "               , conrelid" _
    & Line & "               , confrelid" _
    & Line & "               , contype" _
    & Line & "               , conkey" _
    & Line & "               , confkey" _
    & Line & "               , confupdtype"
    syntax = syntax _
    & Line & "               , confdeltype" _
    & Line & "               , confmatchtype" _
    & Line & "               , generate_subscripts(conkey, 1) as conattind" _
    & Line & "           FROM pg_constraint" _
    & Line & "          WHERE contype = 'f'" _
    & Line & "           ) AS con" _
    & Line & "       ON con.conrelid = c.oid" _
    & Line & "     JOIN pg_attribute a" _
    & Line & "       ON a.attrelid = c.oid" _
    & Line & "      AND a.attnum = con.conkey[conattind]" _
    & Line & "     JOIN pg_class AS cf" _
    & Line & "       ON cf.relnamespace = nsp.oid" _
    & Line & "      AND cf.oid = con.confrelid" _
    & Line & "     JOIN pg_attribute af" _
    & Line & "       ON af.attrelid = c.oid" _
    & Line & "      AND af.attnum = con.confkey[conattind]" _
    & Line & "    WHERE nsp.nspname = current_schema()" _
    & Line & " ORDER BY c.oid" _
    & Line & "        , con.conrelid" _
    & Line & "        , con.conattind"
    
    Dim sSQL                    As String
    sSQL = FormatString(syntax, objTable.TableName)
    
    Dim oRs             As ADODB.Recordset
    Dim curFKID         As String
    Dim objForeignKey   As clsLogicalForeignKey
    
    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    curFKID = ""

    Do While Not oRs.EOF
        '-- For Foreign Key
        If curFKID <> oRs("id") Then
            Set objForeignKey = New clsLogicalForeignKey
            objTable.ForeignKeys.Add objForeignKey

            objForeignKey.RefTableName = oRs("rtName") & ""
            objForeignKey.OnDelete = GetForeignKeyAction(oRs("OnDelete"), "ON DELETE")
            objForeignKey.OnUpdate = GetForeignKeyAction(oRs("OnUpdate"), "ON UPDATE")
            
            curFKID = oRs("id")
        End If

        If Len(objForeignKey.FKcolumns) > 0 Then
            objForeignKey.FKcolumns = objForeignKey.FKcolumns & ", "
        End If
        objForeignKey.FKcolumns = objForeignKey.FKcolumns & oRs("ftColumnName")
        
        If Len(objForeignKey.RefTableColumns) > 0 Then
            objForeignKey.RefTableColumns = objForeignKey.RefTableColumns & ", "
        End If
        objForeignKey.RefTableColumns = objForeignKey.RefTableColumns & oRs("rtColumnName")

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
    
    syntax = "   SELECT c.oid" _
    & Line & "        , a.attnum" _
    & Line & "        , a.attname as column_name" _
    & Line & "        , not a.attnotnull as is_nullable" _
    & Line & "        , t.typname as data_type" _
    & Line & "        , a.atttypmod - 4 as character_maximum_length" _
    & Line & "        , a.atttypmod / 65536  as numeric_precision" _
    & Line & "        , a.atttypmod % 65535 - a.atttypmod / 65536 - 4 as numeric_scale" _
    & Line & "        , ad.adsrc as column_default" _
    & Line & "     FROM pg_namespace AS nsp" _
    & Line & "     JOIN pg_class AS c" _
    & Line & "       ON c.relnamespace = nsp.oid" _
    & Line & "      AND c.relname = '{0:table name}'" _
    & Line & "     JOIN pg_attribute a" _
    & Line & "       ON a.attrelid = c.oid" _
    & Line & "      AND a.attnum > 0" _
    & Line & "     JOIN pg_type t" _
    & Line & "       ON t.oid = a.atttypid" _
    & Line & "LEFT JOIN pg_attrdef ad" _
    & Line & "       ON ad.adrelid = c.oid" _
    & Line & "      AND ad.adnum = a.attnum" _
    & Line & "    WHERE nsp.nspname = current_schema()" _
    & Line & " ORDER BY c.oid" _
    & Line & "        , a.attnum"

    Dim sSQL                    As String
    sSQL = FormatString(syntax, objTable.TableName)
    
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
                                            oRs("character_maximum_length"), _
                                            oRs("numeric_precision"), _
                                            oRs("numeric_scale"))
        objColumn.Nullable = oRs("is_nullable")
        objColumn.Default = oRs("column_default") & ""
        If Not IsNull(oRs("column_comment")) Then
            objColumn.DBDescription = oRs("column_comment")
        Else
            objColumn.DBDescription = ""
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
                        maxLength As Long, _
                        precision As Integer, _
                        type_scale As Integer) As String
    Dim DataType As String

    DataType = LCase(typeName)
    Select Case DataType
        Case "char", "varchar", "binary", "varbinary"
            DataType = DataType & "(" & maxLength & ")"
        Case "nvarchar", "nchar"
            DataType = DataType & "(" & maxLength / 2 & ")"
        Case "numeric"
            If precision > 0 And type_scale >= 0 Then
                DataType = DataType & "(" & precision & "," & type_scale & ")"
            End If
    End Select

    GetColumnDataType = DataType
End Function

Private Function GetForeignKeyAction(action As String, _
            keyType As String)
    '-- a = no action, r = restrict, c = cascade, n = set null, d = set default
    Dim foreignKeyAction As String
    
    Select Case action
    Case "r"
        foreignKeyAction = keyType & " RESTRICT"
    Case "c"
        foreignKeyAction = keyType & " CASCADE"
    Case "n"
        foreignKeyAction = keyType & " SET NULL"
    Case "d"
        foreignKeyAction = keyType & " SET DEFAULT"
    End Select
    
    GetForeignKeyAction = foreignKeyAction
End Function


