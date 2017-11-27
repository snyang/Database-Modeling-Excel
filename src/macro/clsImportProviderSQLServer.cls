VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsImportProviderSQLServer"
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
    Const ProviderOdbc = "SQLOLEDB"
    
    Dim candidateProviders(1 To 1) As String
    candidateProviders(1) = ProviderOdbc
    
    IImportProvider_Providers = candidateProviders
End Property

Public Property Get IImportProvider_SupportSelectDatabase() As Boolean
    IImportProvider_SupportSelectDatabase = True
End Property

Public Function IImportProvider_GetDatabasesSql() As String
    
    Dim syntax As String
    
    syntax = "  SELECT name" _
    & Line & "    FROM sysdatabases" _
    & Line & "ORDER BY name"

    IImportProvider_GetDatabasesSql = syntax
    
End Function

Public Function IImportProvider_GetDatabases(ByVal Server As String, _
                    ByVal UserID As String, _
                    ByVal Password As String) As Collection
    Set IImportProvider_GetDatabases = Nothing
End Function

Public Function IImportProvider_GetTablesSql() As String
    Dim syntax As String
    
    syntax = "  SELECT name" _
    & Line & "    FROM sysobjects" _
    & Line & "   WHERE xtype='U'" _
    & Line & "ORDER BY name"
    
    IImportProvider_GetTablesSql = syntax
End Function

Public Function IImportProvider_GetOptions() As clsImportOptions
    If basImport.mImportOptions Is Nothing Then
        Set basImport.mImportOptions = New clsImportOptions
        basImport.mImportOptions.DataSource = ""
        basImport.mImportOptions.DataSourceTip = "Server Name or IP."
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
    
    Dim connString As String
    If Len(Trim(User)) = 0 Then
        connString = "Provider=SQLOLEDB" _
                & ";Integrated Security=SSPI" _
                & ";initial catalog=" & Trim(Database) _
                & ";Data Source=" & Trim(Server) _
                & ";Connect Timeout=15"
    Else
        connString = "Provider=SQLOLEDB" _
                & ";UID=" & Trim(User) _
                & ";PWD=" & Trim(Password) _
                & ";initial catalog=" & Trim(Database) _
                & ";Data Source=" & Trim(Server) _
                & ";Connect Timeout=15"
    End If
    
    IImportProvider_CreateConnectionString = connString
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
    
    syntax = "     SELECT sysindexes.indid" _
    & Line & "          , IndexName = sysindexes.name" _
    & Line & "          , IsPK = CASE WHEN sysobjects.xtype='PK' THEN 1" _
    & Line & "                   ELSE 0" _
    & Line & "                   End" _
    & Line & "          , IsCluster = INDEXPROPERTY(O.id, sysindexes.[name], 'IsClustered')" _
    & Line & "          , IsUnique = INDEXPROPERTY(O.id, sysindexes.[name], 'IsUnique')" _
    & Line & "          , ColName = syscolumns.[name]" _
    & Line & "       FROM (SELECT id FROM sysobjects WHERE NAME = '{0:table name}') O" _
    & Line & "       JOIN sysindexes" _
    & Line & "         ON sysindexes.id = O.id " _
    & Line & "        AND sysindexes.[name] NOT LIKE '_WA_Sys%'" _
    & Line & "       JOIN sysindexkeys" _
    & Line & "         ON sysindexkeys.id = sysindexes.id" _
    & Line & "        AND sysindexkeys.indid = sysindexes.indid" _
    & Line & "       JOIN syscolumns" _
    & Line & "         ON syscolumns.id = sysindexkeys.id" _
    & Line & "        AND syscolumns.colid = sysindexkeys.colid" _
    & Line & "  LEFT JOIN sysobjects" _
    & Line & "         ON sysobjects.parent_obj = O.id" _
    & Line & "        AND sysobjects.[name] = sysindexes.[name]" _
    & Line & "        AND sysobjects.xtype='PK'" _
    & Line & "   ORDER BY sysindexes.indid" _
    & Line & "          , sysindexkeys.keyno"

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
    If Not oRs Is Nothing Then oRs.Close
    Set oRs = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub RenderForeignKey(conn As ADODB.Connection, objTable As clsLogicalTable)
    Dim syntax As String
    
    syntax = "  select fk.name" _
    & Line & "       , fk.object_id As FKID" _
    & Line & "       , t.name as TableName" _
    & Line & "       , c.name as ColumnName" _
    & Line & "       , ft.name as RefTableName" _
    & Line & "       , fc.name as RefColumnName" _
    & Line & "       , fk.delete_referential_action_desc as DeleteAction" _
    & Line & "       , fk.update_referential_action_desc as UpdateAction" _
    & Line & "    from sys.foreign_keys fk" _
    & Line & "    join sys.foreign_key_columns fkc" _
    & Line & "      on fkc.constraint_object_id = fk.object_id" _
    & Line & "    join sys.tables t" _
    & Line & "      on t.object_id = fkc.parent_object_id" _
    & Line & "    join sys.columns c" _
    & Line & "      on c.object_id = fkc.parent_object_id" _
    & Line & "     and c.column_id = fkc.parent_column_id" _
    & Line & "    join sys.tables ft" _
    & Line & "      on ft.object_id = fkc.referenced_object_id" _
    & Line & "    join sys.columns fc" _
    & Line & "      on fc.object_id = fkc.referenced_object_id" _
    & Line & "     and fc.column_id = fkc.referenced_column_id" _
    & Line & "   where t.name = '{0: table name}'" _
    & Line & "order by t.object_id" _
    & Line & "       , fk.key_index_id" _
    & Line & "       , fkc.constraint_column_id"

    Dim sSQL                    As String
    sSQL = FormatString(syntax, objTable.TableName)
    
    Dim oRs             As ADODB.Recordset
    Dim curFKID         As Long
    Dim objForeignKey   As clsLogicalForeignKey
    
    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    curFKID = 0

    Do While Not oRs.EOF
        '-- For Foreign Key
        If curFKID <> oRs("FKID") Then
            Set objForeignKey = New clsLogicalForeignKey
            objTable.ForeignKeys.Add objForeignKey

            objForeignKey.RefTableName = oRs("RefTableName") & ""
            objForeignKey.OnDelete = GetForeignKeyAction("ON DELETE", oRs("DeleteAction"))
            objForeignKey.OnUpdate = GetForeignKeyAction("ON UPDATE", oRs("UpdateAction"))
            
            curFKID = oRs("FKID")
        End If

        If Len(objForeignKey.FKcolumns) > 0 Then
            objForeignKey.FKcolumns = objForeignKey.FKcolumns & ", "
        End If
        objForeignKey.FKcolumns = objForeignKey.FKcolumns & oRs("ColumnName")
        
        If Len(objForeignKey.RefTableColumns) > 0 Then
            objForeignKey.RefTableColumns = objForeignKey.RefTableColumns & ", "
        End If
        objForeignKey.RefTableColumns = objForeignKey.RefTableColumns & oRs("RefColumnName")

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

Private Function GetForeignKeyAction(ActionName As String, ActionDes As String) As String
    GetForeignKeyAction = ""
    If IsNull(ActionDes) Then
        Exit Function
    End If
    
    ActionDes = UCase(ActionDes)
    If ActionDes = "NO_ACTION" Then
        GetForeignKeyAction = ""
    ElseIf ActionDes = "CASCADE" Then
        GetForeignKeyAction = ActionName & " " & "CASCADE"
    ElseIf ActionDes = "SET_DEFAULT" Then
        GetForeignKeyAction = ActionName & " " & "SET DEFAULT"
    ElseIf ActionDes = "SET_NULL" Then
        GetForeignKeyAction = ActionName & " " & "SET NULL"
    Else
        GetForeignKeyAction = ActionName & " " & Replace(ActionDes, "_", " ")
    End If
End Function

Public Sub RenderColumn(conn As ADODB.Connection, objTable As clsLogicalTable)
    Dim syntax As String
    
    syntax = "    SELECT syscolumns.name" _
    & Line & "         , type_name = systypes.name" _
    & Line & "         , length = syscolumns.length" _
    & Line & "         , [precision] = syscolumns.xprec" _
    & Line & "         , scale =syscolumns.xscale" _
    & Line & "         , is_identity = ISNULL(COLUMNPROPERTY(sysobjects.id, syscolumns.[name], 'IsIdentity'), 0)" _
    & Line & "         , identity_incr = IDENT_INCR(sysobjects.[name])" _
    & Line & "         , identity_seed = IDENT_SEED(sysobjects.[name])" _
    & Line & "         , is_nullable = syscolumns.isnullable" _
    & Line & "         , default_definition = syscomments.text" _
    & Line & "         , description = cd.value" _
    & Line & "         , table_description = td.value" _
    & Line & "      FROM sysobjects" _
    & Line & "      JOIN syscolumns" _
    & Line & "        ON syscolumns.id = sysobjects.id" _
    & Line & " LEFT JOIN syscomments" _
    & Line & "        ON syscomments.id = syscolumns.cdefault" _
    & Line & " LEFT JOIN systypes" _
    & Line & "        ON systypes.xusertype = syscolumns.xusertype"
    syntax = syntax _
    & Line & " LEFT JOIN fn_listextendedproperty('MS_Description', 'user', 'dbo', 'table', '{0:table name}', 'column', NULL) cd" _
    & Line & "        ON cd.objname collate SQL_Latin1_General_CP1_CI_AS = syscolumns.name" _
    & Line & " LEFT JOIN fn_listextendedproperty('MS_Description', 'user', 'dbo', 'table', '{0:table name}', NULL, NULL) td" _
    & Line & "        ON td.objtype = 'TABLE'" _
    & Line & "     WHERE sysobjects.name = '{0:table name}'" _
    & Line & "       AND sysobjects.xtype='U'" _
    & Line & "  ORDER BY syscolumns.colid"

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
        
        objColumn.ColumnName = oRs("name") & ""
        objColumn.DataType = GetColumnDataType( _
                                            oRs("type_name"), _
                                            oRs("length"), _
                                            oRs("precision"), _
                                            oRs("scale"), _
                                            oRs("is_identity"), _
                                            IIf(IsNull(oRs("identity_seed")), 0, oRs("identity_seed")), _
                                            IIf(IsNull(oRs("identity_incr")), 0, oRs("identity_incr")))
        objColumn.Nullable = oRs("is_nullable")
        If Not IsNull(oRs("default_definition")) Then
            objColumn.Default = StripBrackets(oRs("default_definition"))
        Else
            objColumn.Default = ""
        End If
        
        If Not IsNull(oRs("description")) Then
            objColumn.DBDescription = oRs("description")
        End If
        
        objTable.Comment = oRs("table_description") & ""

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
                        maxLength As Integer, _
                        precision As Integer, _
                        type_scale As Integer, _
                        is_identity As Integer, _
                        seed As Integer, _
                        incr As Integer) As String
    Dim DataType As String

    DataType = LCase(typeName)
    Select Case DataType
        Case "char", "varchar", "binary", "varbinary"
            If maxLength < 0 Then
                DataType = DataType & "(max)"
            Else
                DataType = DataType & "(" & maxLength & ")"
            End If
        Case "nvarchar", "nchar"
            If maxLength < 0 Then
                DataType = DataType & "(max)"
            Else
                DataType = DataType & "(" & maxLength / 2 & ")"
            End If
        Case "int", "bigint", "smallint"
            If is_identity Then
                DataType = DataType & " IDENTITY (" & seed & "," & incr & ")"
            End If
        Case "numeric", "decimal"
            If precision < 18 Then
                DataType = DataType & "(" & precision & ", " & type_scale & ")"
            End If
    End Select

    GetColumnDataType = DataType
End Function

