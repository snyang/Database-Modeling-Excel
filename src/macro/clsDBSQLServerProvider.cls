VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBSQLServerProvider"
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
Option Explicit
Implements clsIDatabaseProvider

Const SQL_Statement_End = ";"
Const ForeignKeyNameVariable = "@FkName"

'-----------------------------------------------------------
'-- SQL Server
'-----------------------------------------------------------

Public Sub clsIDatabaseProvider_GetSQLCreateTable(objLogicalTableCollection As Collection, _
                            ByVal withComment As Boolean, _
                            Optional outputOptions As clsOutputOptions)

    Output_Initialize outputOptions
    
    OutputCreateTableSQL objLogicalTableCollection, withComment

    Output_Copy
End Sub

Public Sub clsIDatabaseProvider_GetSQLDropTable(objLogicalTableCollection As Collection, _
                            Optional outputOptions As clsOutputOptions)
    
    Output_Initialize outputOptions
    
    OutputDropTableSQL objLogicalTableCollection
    
    Output_Copy

End Sub

Public Sub clsIDatabaseProvider_GetSQLDropAndCreateTable(objLogicalTableCollection As Collection, _
                            ByVal withComment As Boolean, _
                            Optional outputOptions As clsOutputOptions)
    
    Output_Initialize outputOptions
    
    OutputDropTableSQL objLogicalTableCollection
    
    OutputCreateTableSQL objLogicalTableCollection, withComment
    
    Output_Copy

End Sub

Public Sub clsIDatabaseProvider_GetSQLCreateTableIfNotExists(objLogicalTableCollection As Collection, _
                            Optional outputOptions As clsOutputOptions)
    Dim sSQL                As String
    Dim sqlCreateFK         As String
    Dim objLogicalTable     As clsLogicalTable
    
    Output_Initialize outputOptions
    
    For Each objLogicalTable In objLogicalTableCollection
        Call Get_SQL_CreateTableIfNotExists_Single(objLogicalTable, _
                                            sSQL, _
                                            sqlCreateFK)
        Output_WriteLine sSQL & Line & SQL_Statement_End & Line
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & Line & SQL_Statement_End & Line, 1
        End If
    Next
    
    Output_Copy
End Sub

Public Sub OutputCreateTableSQL(objLogicalTableCollection As Collection, _
                            ByVal withComment As Boolean)
    Dim sSQL                As String
    Dim sqlCreateFK         As String
    Dim objLogicalTable     As clsLogicalTable
    
    '-- Create tables
    For Each objLogicalTable In objLogicalTableCollection
        Call Get_SQL_CreateTable_Single(objLogicalTable, _
                                            withComment, _
                                            sSQL, _
                                            sqlCreateFK)
        Output_WriteLine sSQL & Line & SQL_Statement_End & Line
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & Line & SQL_Statement_End, 1
        End If
    Next

End Sub

Private Sub OutputDropTableSQL(objLogicalTableCollection As Collection)
    Dim syntaxPreRemoveFK      As String
    Dim sSQL            As String
    Dim objLogicalTable As clsLogicalTable
    
    '-- Prepare to drop foreign keys
    syntaxPreRemoveFK = "-- Prepare to remove foreign keys" _
    & Line & "DECLARE {0:@FkName}  SYSNAME"
    sSQL = FormatString(syntaxPreRemoveFK, ForeignKeyNameVariable)
    Output_WriteLine sSQL & Line
    
    '-- Drop foreign key relation
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_RemoveFK_Single(objLogicalTable)
        Output_WriteLine sSQL & Line & SQL_Statement_End & Line
    Next
    
    '-- Drop tables
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_DropTable_Single(objLogicalTable)
        Output_WriteLine sSQL & Line & SQL_Statement_End & Line
    Next
    
End Sub

Private Sub Get_SQL_CreateTableIfNotExists_Single(objLogicalTable As clsLogicalTable, _
                                ByRef sqlCreateTable As String, _
                                ByRef sqlCreateFK As String)
    Dim sSQL            As String
    
    sSQL = "IF NOT EXISTS (" _
        & Line & "  SELECT *" _
        & Line & "    FROM INFORMATION_SCHEMA.TABLES" _
        & Line & "   WHERE TABLE_TYPE = 'BASE TABLE'" _
        & Line & "     AND TABLE_SCHEMA = 'dbo'" _
        & Line & "     AND TABLE_NAME = '{0:table name}'" _
        & Line & "  )" _
        & Line & "BEGIN" _
        & Line & "{1:create table sql}" _
        & Line & "END"
    
    Call Get_SQL_CreateTable_Single(objLogicalTable, _
                                False, _
                                sqlCreateTable, _
                                sqlCreateFK)
                                
    sqlCreateTable = FormatString(sSQL, _
                    objLogicalTable.TableName, _
                    sqlCreateTable)
    
End Sub

Private Function Get_SQL_RemoveFK_Single(objLogicalTable As clsLogicalTable) As String
    Dim syntaxRemoveFK      As String
    
    syntaxRemoveFK = "-- Drop foreign keys for table '{0:table name}'" _
            & Line & "DECLARE fk_cursor CURSOR FOR " _
            & Line & "  SELECT CONSTRAINT_NAME" _
            & Line & "    FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS" _
            & Line & "   WHERE TABLE_SCHEMA = 'dbo'" _
            & Line & "     AND TABLE_NAME = '{0:table name}'" _
            & Line & "     AND CONSTRAINT_TYPE = 'FOREIGN KEY'" _
            & Line & "ORDER BY CONSTRAINT_NAME" _
            & Line & "" _
            & Line & "OPEN fk_cursor" _
            & Line & "FETCH NEXT FROM fk_cursor INTO {1:@FkName}" _
            & Line & "WHILE @@FETCH_STATUS = 0" _
            & Line & "BEGIN" _
            & Line & "  EXEC('ALTER TABLE [{0:table name}] DROP CONSTRAINT ' + {1:@FkName})" _
            & Line & "  FETCH NEXT FROM fk_cursor INTO {1:@FkName}" _
            & Line & "END" _
            & Line & "" _
            & Line & "CLOSE fk_cursor" _
            & Line & "DEALLOCATE fk_cursor"

    '-- Return
    Get_SQL_RemoveFK_Single = FormatString(syntaxRemoveFK, objLogicalTable.TableName, ForeignKeyNameVariable)
End Function

Public Sub Get_SQL_CreateTable_Single(objLogicalTable As clsLogicalTable, _
                                ByVal withComment As Boolean, _
                                ByRef sqlCreateTable As String, _
                                ByRef sqlCreateFK As String)
    Dim objColumn           As clsLogicalColumn
    Dim objIndex            As clsLogicalIndex
    Dim objForeignKey       As clsLogicalForeignKey
    Dim isFirst             As Boolean
    Dim syntaxTable         As String
    Dim syntaxColumn        As String
    Dim syntaxDefault       As String
    Dim syntaxPrimaryKey    As String
    Dim syntaxUniqueIndex   As String
    Dim syntaxNoUniqueIndex As String
    
    Dim sqlComment          As String
    
    syntaxTable = "--  Create table '{0:table name}'" _
                & Line & "CREATE TABLE [{0:table name}] (" _
                & "{1:columns definition}" _
                & "{2:primary key}" _
                & "{3:indexes}" _
                & Line & ")" _
                & "{4:non unique indexes}"
    syntaxColumn = "  {0:(i = 1 ? space : ,)}[{1:column name}] {2:data type} {3:nullable}{4:default}"
    syntaxDefault = " CONSTRAINT {0:default name} DEFAULT ({1:default value})"
    
    Dim sqlColumn           As String
    Dim sqlPrimaryKey       As String
    Dim sqlUniqueIndex      As String
    Dim sqlNoUniqueIndex    As String
    
    '-- Create Columns sql
    sqlColumn = ""
    isFirst = True
    For Each objColumn In objLogicalTable.Columns
        With objColumn
            sqlColumn = sqlColumn _
                    & Line & FormatString(syntaxColumn, _
                                IIf(isFirst, " ", ","), _
                                .ColumnName, _
                                .DataType, _
                                IIf(.Nullable, "NULL", "NOT NULL"), _
                                FormatString(IIf(Len(.Default) = 0, "", syntaxDefault), _
                                    SQL_Render_DF_Name(objLogicalTable, objColumn), _
                                    StripBrackets(.Default, True)))
        End With
        isFirst = False
    Next
    
    syntaxPrimaryKey = "  ,CONSTRAINT {0:primary key name} PRIMARY KEY {1:clustered tag} ({2:columns})"
    syntaxUniqueIndex = "  ,CONSTRAINT {0:index name} UNIQUE{1:clustered tag} {2:columns with bracket}"
    syntaxNoUniqueIndex = "CREATE{1:clustered tag} INDEX {0:index name} ON [{2:table name}] {3:columns}"
    '-- Primary key SQL
    sqlPrimaryKey = ""
    With objLogicalTable.PrimaryKey
        If Len(.PKcolumns) > 0 Then
            sqlPrimaryKey = Line & FormatString(syntaxPrimaryKey, _
                                SQL_Render_PK_Name(objLogicalTable), _
                                IIf(.IsClustered, "CLUSTERED", "NONCLUSTERED"), _
                                 .PKcolumns)
        End If
    End With
    
    '-- Unqiue Indexes SQL
    sqlUniqueIndex = ""
    For Each objIndex In objLogicalTable.Indexes
        With objIndex
            If .IsUnique Then
                sqlUniqueIndex = sqlUniqueIndex _
                        & Line & FormatString(syntaxUniqueIndex, _
                                    SQL_Render_IK_Name(objLogicalTable, objIndex), _
                                    IIf(.IsClustered, " CLUSTERED", ""), _
                                    .IKColumns)
            End If
        End With
        isFirst = False
    Next

    '-- non-unique Indexes SQL
    sqlNoUniqueIndex = ""
    For Each objIndex In objLogicalTable.Indexes
        With objIndex
            If Not .IsUnique Then
                sqlNoUniqueIndex = sqlNoUniqueIndex _
                        & Line & FormatString(syntaxNoUniqueIndex, _
                                    SQL_Render_IK_Name(objLogicalTable, objIndex), _
                                    IIf(.IsClustered, " CLUSTERED", ""), _
                                    objLogicalTable.TableName, _
                                    .IKColumns)
            End If
        End With
    Next
    
    '-- Generate table sql
    sqlCreateTable = FormatString(syntaxTable, _
                        objLogicalTable.TableName, _
                        sqlColumn, _
                        sqlPrimaryKey, _
                        sqlUniqueIndex, _
                        sqlNoUniqueIndex)
    
    '-- Generate Foreign key SQL
    Dim sqlFKHead       As String
    Dim sqlFK           As String
    sqlFKHead = "-- Foreign keys for table '{0:table name}'" _
        & Line & "ALTER TABLE [{0:Table Name}] ADD"
    sqlFK = "{0:(i=1? :,)} CONSTRAINT {1:foreign key name}" _
    & Line & "  FOREIGN KEY ({2:column[,..n]})" _
    & Line & "  REFERENCES {3:ref_info}"
    
    sqlCreateFK = ""
    If objLogicalTable.ForeignKeys.Count > 0 Then
        sqlCreateFK = FormatString(sqlFKHead, objLogicalTable.TableName)
        isFirst = True
        For Each objForeignKey In objLogicalTable.ForeignKeys
            With objForeignKey
                sqlCreateFK = sqlCreateFK _
                                & Line & FormatString(sqlFK, _
                                            IIf(isFirst, " ", ","), _
                                            SQL_Render_FK_Name(objLogicalTable, objForeignKey), _
                                            .FKcolumns, _
                                            .RefTableAndColumns & IIf(Len(.fkOption) = 0, "", " " & .fkOption))
            End With
            isFirst = False
        Next
    End If
    
    '-- Column comment
    Dim syntaxTableComment   As String
    syntaxTableComment = "EXECUTE sp_addextendedproperty N'MS_Description'" _
    & Line & "  , N{0:column comment}" _
    & Line & "  , N'user', N'dbo', N'table', N'{1:table name}'"
    Dim syntaxColumnComment   As String
    syntaxColumnComment = "EXECUTE sp_addextendedproperty N'MS_Description'" _
    & Line & "  , N{0:column comment}" _
    & Line & "  , N'user', N'dbo', N'table', N'{1:table name}', N'column', N'{2:column name}'"
    sqlComment = ""
    
    If withComment Then
        '-- For table description
        If Len(objLogicalTable.Comment) Then
            sqlComment = sqlComment _
                    & Line & FormatString(syntaxTableComment, _
                        SQL_ToSQL(objLogicalTable.Comment), _
                        objLogicalTable.TableName)
        End If
        
        '-- For column description
        For Each objColumn In objLogicalTable.Columns
            With objColumn
                sqlComment = sqlComment _
                        & Line & FormatString(syntaxColumnComment, _
                            SQL_ToSQL(.DBDescription), _
                            objLogicalTable.TableName, _
                            .ColumnName)
            End With
        Next
        
        If Len(sqlComment) Then
            sqlCreateTable = sqlCreateTable _
            & Line & sqlComment
        End If
    End If
    
End Sub

Private Function Get_SQL_DropTable_Single(objLogicalTable As clsLogicalTable) As String
    Dim sSQL            As String
    
    sSQL = "-- Drop table '{0:table name}'" _
        & Line & "IF EXISTS (" _
        & Line & "  SELECT * FROM INFORMATION_SCHEMA.TABLES" _
        & Line & "  WHERE TABLE_TYPE = 'BASE TABLE'" _
        & Line & "  AND TABLE_SCHEMA = 'dbo'" _
        & Line & "  AND TABLE_NAME = '{0:table name}'" _
        & Line & "  )" _
        & Line & "BEGIN" _
        & Line & "  DROP TABLE [dbo].[{0:table name}]" _
        & Line & "END"
   
    '-- Return
    Get_SQL_DropTable_Single = FormatString(sSQL, _
                                objLogicalTable.TableName)
End Function

Private Function SQL_Render_TableName(ByVal TableName As String) As String
    SQL_Render_TableName = Replace(Replace(TableName, " ", ""), "_", "")
End Function

Private Function SQL_Render_PK_Name(objLogicalTable As clsLogicalTable) As String
    SQL_Render_PK_Name = "PK_" & SQL_Render_TableName(objLogicalTable.TableName)
End Function

Private Function SQL_Render_FK_Name(objLogicalTable As clsLogicalTable, _
                                objLogicalForeignKey As clsLogicalForeignKey) As String
    SQL_Render_FK_Name = "FK_" & SQL_Render_TableName(objLogicalTable.TableName) _
                            & "_" & Replace(Replace(Replace(objLogicalForeignKey.FKcolumns, " ", ""), "_", ""), ",", "_")
End Function

Private Function SQL_Render_IK_Name(objLogicalTable As clsLogicalTable, _
                                objLogicalIndex As clsLogicalIndex) As String
    SQL_Render_IK_Name = "IK_" & SQL_Render_TableName(objLogicalTable.TableName) _
                            & "_" & Replace(Replace(Replace(Replace(Replace(objLogicalIndex.IKColumns, _
                                                                    " ", ""), _
                                                            "(", ""), _
                                                    ")", ""), _
                                            "_", ""), _
                                    ",", "_")
End Function

Private Function SQL_Render_DF_Name(objLogicalTable As clsLogicalTable, _
                                objLogicalColumn As clsLogicalColumn) As String
    SQL_Render_DF_Name = "DF_" & SQL_Render_TableName(objLogicalTable.TableName) & "_" & objLogicalColumn.ColumnName
End Function
