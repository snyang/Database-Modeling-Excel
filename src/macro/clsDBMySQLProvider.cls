VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBMySQLProvider"
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

'-----------------------------------------------------------
'-- MySQL
'-----------------------------------------------------------
Private Const SP_REMVOE_FK_NAME = "__tmp_removeFK"
'-------------------------------------

Public Sub clsIDatabaseProvider_GetSQLCreateTable(objLogicalTableCollection As Collection, _
                            ByVal withComment As Boolean, _
                            Optional outputOptions As clsOutputOptions)
    Dim sSQL                As String
    Dim sqlCreateFK         As String
    Dim sqlAddComment       As String
    Dim objLogicalTable     As clsLogicalTable
    
    Output_Initialize outputOptions
    
    '-- Create tables
    For Each objLogicalTable In objLogicalTableCollection
        Call Get_SQL_CreateTable_Single(objLogicalTable, _
                                            withComment, _
                                            sSQL, _
                                            sqlCreateFK, _
                                            False)
        Output_WriteLine sSQL & Line
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & Line, 1
        End If
    Next
    
    Output_Copy
End Sub

Public Sub clsIDatabaseProvider_GetSQLDropTable(objLogicalTableCollection As Collection, _
                            Optional outputOptions As clsOutputOptions)
    Dim sSQL            As String
    Dim objLogicalTable As clsLogicalTable
    
    Output_Initialize outputOptions
    
    '-- Create the store procedue of drop foreign key
    sSQL = Get_SQL_Pre_RemoveFK()
    Output_WriteLine sSQL & Line
    
    '-- Drop foreign key relation
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_RemoveFK_Single(objLogicalTable)
        Output_WriteLine sSQL & Line
    Next
    
    '-- Drop the store procedue of drop foreign key
    sSQL = Get_SQL_Post_RemoveFK()
    Output_WriteLine sSQL & Line
    
    '-- Drop tables
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_DropTable_Single(objLogicalTable)
        Output_WriteLine sSQL & Line
    Next
    
    Output_Copy
End Sub

Public Sub clsIDatabaseProvider_GetSQLDropAndCreateTable(objLogicalTableCollection As Collection, _
                            ByVal withComment As Boolean, _
                            Optional outputOptions As clsOutputOptions)
    Dim sSQL                As String
    Dim sqlCreateFK         As String
    Dim objLogicalTable     As clsLogicalTable
    
    Output_Initialize outputOptions
    
    '-- Create the store procedue of drop foreign key
    sSQL = Get_SQL_Pre_RemoveFK()
    Output_WriteLine sSQL & Line
    
    '-- Drop foreign key relation
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_RemoveFK_Single(objLogicalTable)
        Output_WriteLine sSQL & Line
    Next
    
    '-- Drop the store procedue of drop foreign key
    sSQL = Get_SQL_Post_RemoveFK()
    Output_WriteLine sSQL & Line
    
    '-- Drop tables
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_DropTable_Single(objLogicalTable)
        Output_WriteLine sSQL & Line
    Next
    
    '-- Create tables
    For Each objLogicalTable In objLogicalTableCollection
        Call Get_SQL_CreateTable_Single(objLogicalTable, _
                                            withComment, _
                                            sSQL, _
                                            sqlCreateFK, _
                                            False)
        Output_WriteLine sSQL & Line
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & Line, 1
        End If
    Next
    
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
        Output_WriteLine sSQL & Line
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & Line, 1
        End If
    Next
    
    Output_Copy
End Sub

Private Sub Get_SQL_CreateTableIfNotExists_Single(objLogicalTable As clsLogicalTable, _
                                ByRef sqlCreateTable As String, _
                                ByRef sqlCreateFK As String)
    Call Get_SQL_CreateTable_Single(objLogicalTable, _
                                False, _
                                sqlCreateTable, _
                                sqlCreateFK, _
                                True)
End Sub

Private Function Get_SQL_RemoveFK_Single(objLogicalTable As clsLogicalTable) As String
    Dim syntaxRemoveFK      As String
    
    syntaxRemoveFK = "-- Remove {0:table name} foreign key constraint" _
            & Line & "CALL {1:sp name for remove FK}('{0:table name}');" _
            
    '-- Return
    Get_SQL_RemoveFK_Single = FormatString(syntaxRemoveFK, objLogicalTable.TableName, SP_REMVOE_FK_NAME)
End Function

Private Function Get_SQL_Pre_RemoveFK() As String
    Dim sSQL        As String
    
    sSQL = "-- temporary store procedue for remove foreign key" _
            & Line & "DROP PROCEDURE IF EXISTS `{0:sp name for remove FK}`;" _
            & Line & "DELIMITER $$" _
            & Line & "CREATE PROCEDURE `{0:sp name for remove FK}` (tableName varchar(64))" _
            & Line & "BEGIN" _
            & Line & "  DECLARE fkName varchar(64);" _
            & Line & "  DECLARE sqlDropFK varchar(250);" _
            & Line & "  DECLARE done INT DEFAULT 0;" _
            & Line & "" _
            & Line & "  DECLARE fkCursor CURSOR FOR" _
            & Line & "    SELECT CONSTRAINT_NAME FROM information_schema.TABLE_CONSTRAINTS TC" _
            & Line & "    WHERE TC.TABLE_SCHEMA = database()" _
            & Line & "    AND   TC.TABLE_NAME = tableName" _
            & Line & "    AND   TC.CONSTRAINT_TYPE = 'FOREIGN KEY';" _
            & Line & "  DECLARE CONTINUE HANDLER FOR SQLSTATE '02000' SET done = 1;" _
            & Line & "" _
            & Line & "  OPEN fkCursor;" _
            & Line & "  FETCH fkCursor INTO fkName;" _
            & Line & "" _
            & Line & "  WHILE done = 0 DO" _
            & Line & "    SET @sqlDropFK = CONCAT('ALTER TABLE `', tableName ,'` DROP FOREIGN KEY `', fkName, '`;');" _
            & Line & "    PREPARE stmt_dropFK FROM @sqlDropFK;" _
            & Line & "    EXECUTE stmt_dropFK;" _
            & Line & "    DEALLOCATE PREPARE stmt_dropFK;"
    sSQL = sSQL & Line & "" _
            & Line & "    FETCH fkCursor INTO fkName;" _
            & Line & "  END WHILE;" _
            & Line & "" _
            & Line & "  CLOSE fkCursor;" _
            & Line & "END $$" _
            & Line & "" _
            & Line & "DELIMITER ;"
    '-- Return
    Get_SQL_Pre_RemoveFK = FormatString(sSQL, SP_REMVOE_FK_NAME)
End Function

Private Function Get_SQL_Post_RemoveFK() As String
    Dim sSQL        As String
    
    sSQL = "-- Remove temporary store procedue for remove foreign key" _
            & Line & "DROP PROCEDURE IF EXISTS `{0:sp name for remove FK}`;" _
            
    '-- Return
    Get_SQL_Post_RemoveFK = FormatString(sSQL, SP_REMVOE_FK_NAME)
End Function

Public Sub Get_SQL_CreateTable_Single(objLogicalTable As clsLogicalTable, _
                                ByVal withComment As Boolean, _
                                ByRef sqlCreateTable As String, _
                                ByRef sqlCreateFK As String, _
                                ByVal IfNotExists As Boolean)
    Dim i               As Integer
    
    Dim syntaxTable         As String
    Dim syntaxColumn        As String
    Dim syntaxDefault       As String
    Dim syntaxPrimaryKey    As String
    Dim syntaxUniqueIndex   As String
    Dim syntaxNoUniqueIndex As String
    
    syntaxTable = "CREATE TABLE {5:if not exists}`{0:table name}` (" _
                & "{1:columns definition}" _
                & "{2:primary key}" _
                & "{3:indexes}" _
                & Line & ");" _
                & "{4:non unique indexes}"
    syntaxColumn = "  {0:(i = 1 ? space : ,)}`{1:column name}` {2:data type} {3:nullable} {4:default} {5:comment}"
    syntaxDefault = " DEFAULT {1:default value}"
    syntaxPrimaryKey = "  ,CONSTRAINT {0:primary key name} PRIMARY KEY ({1:columns})"
    syntaxUniqueIndex = "  ,CONSTRAINT {0:index name} UNIQUE {1:columns with bracket}"
    syntaxNoUniqueIndex = "CREATE INDEX {0:index name} ON `{1:table name}` {2:columns};"
    
    Dim sqlColumn           As String
    Dim sqlPrimaryKey       As String
    Dim sqlUniqueIndex      As String
    Dim sqlNoUniqueIndex    As String
    
    '-- Create Columns sql
    sqlColumn = ""
    For i = 1 To objLogicalTable.Columns.Count
        With objLogicalTable.Columns(i)
            sqlColumn = sqlColumn _
                    & Line & FormatString(syntaxColumn, _
                                IIf(i = 1, " ", ","), _
                                .ColumnName, _
                                .DataType, _
                                IIf(.Nullable, "NULL", "NOT NULL"), _
                                FormatString(IIf(Len(.Default) = 0, "", syntaxDefault), _
                                    SQL_Render_DF_Name(objLogicalTable, objLogicalTable.Columns(i)), _
                                    .Default), _
                                IIf(withComment, _
                                    "COMMENT " & SQL_ToSQL(IIf(Len(.Comment) = 0, .ColumnLabel, .Comment)), _
                                    ""))
        End With
    Next
    
    '-- Primary key SQL
    sqlPrimaryKey = ""
    With objLogicalTable.PrimaryKey
        If Len(.PKcolumns) > 0 Then
            sqlPrimaryKey = Line & FormatString(syntaxPrimaryKey, _
                                SQL_Render_PK_Name(objLogicalTable), _
                                .PKcolumns)
        End If
    End With
    
    '-- Unqiue Indexes SQL
    sqlUniqueIndex = ""
    For i = 1 To objLogicalTable.Indexes.Count
        With objLogicalTable.Indexes(i)
            If .IsUnique Then
                sqlUniqueIndex = sqlUniqueIndex _
                        & Line & FormatString(syntaxUniqueIndex, _
                                    SQL_Render_IK_Name(objLogicalTable, objLogicalTable.Indexes(i)), _
                                    .IKColumns)
            End If
        End With
    Next

    '-- non-unique Indexes SQL
    sqlNoUniqueIndex = ""
    For i = 1 To objLogicalTable.Indexes.Count
        With objLogicalTable.Indexes(i)
            If Not .IsUnique Then
                sqlNoUniqueIndex = sqlNoUniqueIndex _
                        & Line _
                        & FormatString(syntaxNoUniqueIndex, _
                            SQL_Render_IK_Name(objLogicalTable, objLogicalTable.Indexes(i)), _
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
                        sqlNoUniqueIndex, _
                        IIf(IfNotExists, "IF NOT EXISTS ", ""))
    
    '-- Generate Foreign key SQL
    Dim sqlFK           As String
    sqlFK = "ALTER TABLE `{0:Table Name}` ADD CONSTRAINT {1:foreign key name} FOREIGN KEY ({2:column[,..n]}) REFERENCES {3:ref_info};"
    
    sqlCreateFK = ""
    If objLogicalTable.ForeignKeys.Count > 0 Then
        For i = 1 To objLogicalTable.ForeignKeys.Count
            With objLogicalTable.ForeignKeys(i)
                sqlCreateFK = sqlCreateFK _
                                & Line _
                                & FormatString(sqlFK, _
                                    objLogicalTable.TableName, _
                                    SQL_Render_FK_Name(objLogicalTable, objLogicalTable.ForeignKeys(i)), _
                                    .FKcolumns, _
                                    .RefTableAndColumns & IIf(Len(.fkOption) = 0, "", " " & .fkOption))
            End With
        Next
    End If
End Sub

Private Function Get_SQL_DropTable_Single(objLogicalTable As clsLogicalTable) As String
    Dim sSQL            As String
    
    sSQL = "DROP TABLE IF EXISTS `{0:table name}`;"
   
    '-- Return
    Get_SQL_DropTable_Single = FormatString(sSQL, _
                                objLogicalTable.TableName, _
                                objLogicalTable.Comment)
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


