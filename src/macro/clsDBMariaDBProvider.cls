VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBMariaDBProvider"
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
Implements clsIDatabaseProvider

'-----------------------------------------------------------
'-- MariaDB
'-----------------------------------------------------------
Private Const SP_Drop_FK_Name = "__tmp_dropFK"
'-------------------------------------

Public Sub clsIDatabaseProvider_GetSQLCreateTable(objTableCollection As Collection, _
                            ByVal withComment As Boolean, _
                            Optional outputOptions As clsOutputOptions)
    
    Output_Initialize outputOptions
    
    OutputCreateTableSQL objTableCollection, withComment
    
    Output_Copy
End Sub

Public Sub clsIDatabaseProvider_GetSQLDropTable(objTableCollection As Collection, _
                            Optional outputOptions As clsOutputOptions)
   
    Output_Initialize outputOptions
    
    OutputDropTableSQL objTableCollection
    
    Output_Copy
End Sub

Public Sub clsIDatabaseProvider_GetSQLDropAndCreateTable(objTableCollection As Collection, _
                            ByVal withComment As Boolean, _
                            Optional outputOptions As clsOutputOptions)
    
    Output_Initialize outputOptions
    
    OutputDropTableSQL objTableCollection
    
    OutputCreateTableSQL objTableCollection, withComment
    
    Output_Copy
End Sub

Public Sub clsIDatabaseProvider_GetSQLCreateTableIfNotExists(objTableCollection As Collection, _
                            Optional outputOptions As clsOutputOptions)
    Dim sSQL                As String
    Dim sqlCreateFK         As String
    Dim objTable     As clsLogicalTable
    
    Output_Initialize outputOptions
    
    For Each objTable In objTableCollection
        Call Get_SQL_CreateTableIfNotExists_Single(objTable, _
                                            sSQL, _
                                            sqlCreateFK)
        Output_WriteLine sSQL & Line
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & Line, 1
        End If
    Next
    
    Output_Copy
End Sub

Private Sub OutputCreateTableSQL(objTableCollection As Collection, _
                            ByVal withComment As Boolean)
    Dim sSQL                As String
    Dim sqlCreateFK         As String
    Dim objTable     As clsLogicalTable
    
    '-- Create tables
    For Each objTable In objTableCollection
        Call Get_SQL_CreateTable_Single(objTable, _
                                            withComment, _
                                            sSQL, _
                                            sqlCreateFK, _
                                            False)
        Output_WriteLine sSQL & Line
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & Line, 1
        End If
    Next
End Sub

Private Sub OutputDropTableSQL(objTableCollection As Collection)
    Dim sSQL            As String
    Dim objTable As clsLogicalTable
    
    '-- Create the store procedure of drop foreign key
    sSQL = Get_SQL_Pre_RemoveFK()
    Output_WriteLine sSQL & Line
    
    '-- Drop foreign key relation
    For Each objTable In objTableCollection
        sSQL = Get_SQL_RemoveFK_Single(objTable)
        Output_WriteLine sSQL & Line
    Next
    
    '-- Drop the store procedure of drop foreign key
    sSQL = Get_SQL_Post_RemoveFK()
    Output_WriteLine sSQL & Line
    
    '-- Drop tables
    For Each objTable In objTableCollection
        sSQL = Get_SQL_DropTable_Single(objTable)
        Output_WriteLine sSQL & Line
    Next
End Sub


Private Sub Get_SQL_CreateTableIfNotExists_Single(objTable As clsLogicalTable, _
                                ByRef sqlCreateTable As String, _
                                ByRef sqlCreateFK As String)
    Call Get_SQL_CreateTable_Single(objTable, _
                                False, _
                                sqlCreateTable, _
                                sqlCreateFK, _
                                True)
End Sub

Private Function Get_SQL_RemoveFK_Single(objTable As clsLogicalTable) As String
    Dim syntaxRemoveFK      As String
    
    syntaxRemoveFK = "-- Drop foreign keys for table '{0:table name}'" _
            & Line & "CALL {1:sp name for remove FK}('{0:table name}');" _
            
    '-- Return
    Get_SQL_RemoveFK_Single = FormatString(syntaxRemoveFK, objTable.TableName, SP_Drop_FK_Name)
End Function

Private Function Get_SQL_Pre_RemoveFK() As String
    Dim sSQL        As String
    
    sSQL = "-- Prepare a temporary store procedure to drop foreign keys" _
            & Line & "DROP PROCEDURE IF EXISTS {0:sp name for remove FK};" _
            & Line & "DELIMITER $$" _
            & Line & "CREATE PROCEDURE {0:sp name for remove FK} (tableName varchar(64))" _
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
            & Line & "  WHILE done = 0 DO" _
            & Line & "    SET @sqlDropFK = CONCAT('ALTER TABLE ', tableName ,' DROP FOREIGN KEY ', fkName, ';');" _
            & Line & "    PREPARE stmt_dropFK FROM @sqlDropFK;" _
            & Line & "    EXECUTE stmt_dropFK;" _
            & Line & "    DEALLOCATE PREPARE stmt_dropFK;"
    sSQL = sSQL & Line & "" _
            & Line & "    FETCH fkCursor INTO fkName;" _
            & Line & "  END WHILE;" _
            & Line & "" _
            & Line & "  CLOSE fkCursor;" _
            & Line & "END $$" _
            & Line & "DELIMITER ;"
    '-- Return
    Get_SQL_Pre_RemoveFK = FormatString(sSQL, SP_Drop_FK_Name)
End Function

Private Function Get_SQL_Post_RemoveFK() As String
    Dim sSQL        As String
    
    sSQL = "-- Drop the temporary store procedure of dropping foreign keys" _
            & Line & "DROP PROCEDURE IF EXISTS {0:sp name for drop foreign keys};" _
            
    '-- Return
    Get_SQL_Post_RemoveFK = FormatString(sSQL, SP_Drop_FK_Name)
End Function

Public Sub Get_SQL_CreateTable_Single(objTable As clsLogicalTable, _
                                ByVal withComment As Boolean, _
                                ByRef sqlCreateTable As String, _
                                ByRef sqlCreateFK As String, _
                                ByVal IfNotExists As Boolean)
    Dim objColumn           As clsLogicalColumn
    Dim objFk               As clsLogicalForeignKey
    Dim objIndex            As clsLogicalIndex
    Dim isFirst             As Boolean
    
    Dim syntaxTable         As String
    Dim syntaxColumn        As String
    Dim syntaxTabComment    As String
    Dim syntaxDefault       As String
    Dim syntaxPrimaryKey    As String
    Dim syntaxUniqueIndex   As String
    Dim syntaxNoUniqueIndex As String
    
    syntaxTable = "-- Create table '{0:Table Name}'" _
                & Line & "CREATE TABLE {6:if not exists}{0:table name} (" _
                & "{1:columns definition}" _
                & "{2:primary key}" _
                & "{3:indexes}" _
                & Line & ")" _
                & "{5:table options}" _
                & Line & ";" _
                & "{4:non unique indexes}"
    syntaxColumn = "  {0:(i = 1 ? space : ,)}{1:column name} {2:data type}{3:nullable}{4:default}{5:comment}"
    syntaxDefault = " DEFAULT {1:default value}"
    syntaxTabComment = Line & "   COMMENT = {0:comment}"
    syntaxPrimaryKey = "  ,CONSTRAINT {0:primary key name} PRIMARY KEY ({1:columns})"
    syntaxUniqueIndex = "  ,CONSTRAINT {0:index name} UNIQUE {1:columns with bracket}"
    syntaxNoUniqueIndex = "CREATE INDEX {0:index name} ON {1:table name} {2:columns};"
    
    Dim sqlColumn           As String
    Dim sqlPrimaryKey       As String
    Dim sqlUniqueIndex      As String
    Dim sqlNoUniqueIndex    As String
    
    '-- Create Columns sql
    sqlColumn = ""
    isFirst = True
    For Each objColumn In objTable.Columns
        With objColumn
            sqlColumn = sqlColumn _
                    & Line & FormatString(syntaxColumn, _
                                IIf(isFirst, " ", ","), _
                                .ColumnName, _
                                .DataType, _
                                IIf(.Nullable, " NULL", " NOT NULL"), _
                                FormatString(IIf(Len(.Default) = 0, "", syntaxDefault), _
                                    SQL_Render_DF_Name(objTable, objColumn), _
                                    .Default), _
                                IIf(withComment, _
                                    " COMMENT " & SQL_ToSQL(.DBDescription), _
                                    ""))
        End With
        isFirst = False
    Next
    
    '-- Primary key SQL
    sqlPrimaryKey = ""
    With objTable.PrimaryKey
        If Len(.PKcolumns) > 0 Then
            sqlPrimaryKey = Line & FormatString(syntaxPrimaryKey, _
                                SQL_Render_PK_Name(objTable), _
                                .PKcolumns)
        End If
    End With
    
    '-- Unqiue Indexes SQL
    sqlUniqueIndex = ""
    For Each objIndex In objTable.Indexes
        With objIndex
            If .IsUnique Then
                sqlUniqueIndex = sqlUniqueIndex _
                        & Line & FormatString(syntaxUniqueIndex, _
                                    SQL_Render_IK_Name(objTable, objIndex), _
                                    .IKColumns)
            End If
        End With
    Next

    '-- non-unique Indexes SQL
    sqlNoUniqueIndex = ""
    For Each objIndex In objTable.Indexes
        With objIndex
            If Not .IsUnique Then
                sqlNoUniqueIndex = sqlNoUniqueIndex _
                        & Line _
                        & FormatString(syntaxNoUniqueIndex, _
                            SQL_Render_IK_Name(objTable, objIndex), _
                            objTable.TableName, _
                            .IKColumns)
            End If
        End With
    Next
    
    '-- Generate table sql
    sqlCreateTable = FormatString(syntaxTable, _
                        objTable.TableName, _
                        sqlColumn, _
                        sqlPrimaryKey, _
                        sqlUniqueIndex, _
                        sqlNoUniqueIndex, _
                        IIf(withComment, FormatString(syntaxTabComment, SQL_ToSQL(objTable.Comment)), ""), _
                        IIf(IfNotExists, "IF NOT EXISTS ", ""))
    
    '-- Generate Foreign key SQL
    Dim sqlFK           As String
    Dim sqlPreFK           As String
    sqlFK = "ALTER TABLE {0:Table Name}" _
    & Line & "  ADD CONSTRAINT {1:foreign key name}" _
    & Line & "  FOREIGN KEY ({2:column[,..n]})" _
    & Line & "  REFERENCES {3:ref_info};"
    sqlPreFK = "-- Foreign keys for table '{0:Table Name}'"
    
    sqlCreateFK = ""
    If objTable.ForeignKeys.Count > 0 Then
        sqlCreateFK = FormatString(sqlPreFK, objTable.TableName)
    End If
    For Each objFk In objTable.ForeignKeys
        With objFk
            sqlCreateFK = sqlCreateFK _
                            & Line _
                            & FormatString(sqlFK, _
                                objTable.TableName, _
                                SQL_Render_FK_Name(objTable, objFk), _
                                .FKcolumns, _
                                .RefTableAndColumns & IIf(Len(.fkOption) = 0, "", " " & .fkOption))
        End With
    Next
    
End Sub

Private Function Get_SQL_DropTable_Single(objTable As clsLogicalTable) As String
    Dim sSQL            As String
    
    sSQL = "-- Drop table '{0:table name}'" _
    & Line & "DROP TABLE IF EXISTS {0:table name};"
   
    '-- Return
    Get_SQL_DropTable_Single = FormatString(sSQL, _
                                objTable.TableName, _
                                objTable.Comment)
End Function

Private Function SQL_Render_TableName(ByVal TableName As String) As String
    SQL_Render_TableName = Replace(Replace(TableName, " ", ""), "_", "")
End Function

Private Function SQL_Render_PK_Name(objTable As clsLogicalTable) As String
    SQL_Render_PK_Name = "PK_" & SQL_Render_TableName(objTable.TableName)
End Function

Private Function SQL_Render_FK_Name(objTable As clsLogicalTable, _
                                objLogicalForeignKey As clsLogicalForeignKey) As String
    SQL_Render_FK_Name = "FK_" & SQL_Render_TableName(objTable.TableName) _
                            & "_" & Replace(Replace(Replace(objLogicalForeignKey.FKcolumns, " ", ""), "_", ""), ",", "_")
End Function

Private Function SQL_Render_IK_Name(objTable As clsLogicalTable, _
                                objLogicalIndex As clsLogicalIndex) As String
    SQL_Render_IK_Name = "IK_" & SQL_Render_TableName(objTable.TableName) _
                            & "_" & Replace(Replace(Replace(Replace(Replace(objLogicalIndex.IKColumns, _
                                                                    " ", ""), _
                                                            "(", ""), _
                                                    ")", ""), _
                                            "_", ""), _
                                    ",", "_")
End Function

Private Function SQL_Render_DF_Name(objTable As clsLogicalTable, _
                                objLogicalColumn As clsLogicalColumn) As String
    SQL_Render_DF_Name = "DF_" & SQL_Render_TableName(objTable.TableName) & "_" & objLogicalColumn.ColumnName
End Function
