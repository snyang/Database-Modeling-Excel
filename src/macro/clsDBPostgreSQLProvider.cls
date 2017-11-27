VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBPostgreSQLProvider"
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
'-- PostgreSQL
'Syntax:
'  Comment: -- or /* */
'  String literal modifier: '
'  Identifier modifier: "
' Has concept of inherits
' CREATE LANGUAGE plpgsql;
'
'CREATE INDEX ik_1
'  ON sampletable2
'  USING btree
'  (subitemid);
'ALTER TABLE sampletable2 CLUSTER ON ik_1;
'
'-- Index: ik_2
'
'-- DROP INDEX ik_2;
'
'CREATE UNIQUE INDEX ik_2
'  ON sampletable2
'  USING btree
'  (subitemid);
'-----------------------------------------------------------

Const SQL_Statement_End = ";"
Const IdentifierStart = "" '-- do not use ", otherwise " will be part of table name.
Const IdentifierEnd = ""
Const OutputID_FirstSection = 1
Const OutputID_SecondSection = 2

Public Sub clsIDatabaseProvider_GetSQLCreateTable(objLogicalTableCollection As Collection, _
                            ByVal withComment As Boolean, _
                            Optional outputOptions As clsOutputOptions)
    
    Output_Initialize outputOptions
    
    Call GetSQLCreateTableInternal(objLogicalTableCollection, withComment, False)
    
    Output_Copy
End Sub

Public Sub clsIDatabaseProvider_GetSQLDropTable(objLogicalTableCollection As Collection, _
                            Optional outputOptions As clsOutputOptions)
    Output_Initialize outputOptions
    
    '-- Drop table SQL
    Call GetSQLDropTableInternal(objLogicalTableCollection)

    Output_Copy
End Sub

Public Sub clsIDatabaseProvider_GetSQLDropAndCreateTable(objLogicalTableCollection As Collection, _
                            ByVal withComment As Boolean, _
                            Optional outputOptions As clsOutputOptions)
    Output_Initialize outputOptions
    
    '-- Drop table SQL
    Call GetSQLDropTableInternal(objLogicalTableCollection)
    
    '-- Create tables SQL
    Call GetSQLCreateTableInternal(objLogicalTableCollection, withComment, False)
    
    Output_Copy
End Sub

Public Sub clsIDatabaseProvider_GetSQLCreateTableIfNotExists(objLogicalTableCollection As Collection, _
                            Optional outputOptions As clsOutputOptions)
    Output_Initialize outputOptions
    
    Call GetSQLCreateTableInternal(objLogicalTableCollection, False, True)
    
    Output_Copy
End Sub

Private Sub GetSQLCreateTableInternal(objLogicalTableCollection As Collection, _
                            ByVal withComment As Boolean, _
                            ByVal withIfNotExists As Boolean)
    
    Dim sSQL                As String
    Dim sqlCreateFK         As String
    Dim sqlAddComment       As String
    Dim objLogicalTable     As clsLogicalTable
    '-- Create tables
    For Each objLogicalTable In objLogicalTableCollection
        Call Get_SQL_CreateTable_Single(objLogicalTable, _
                                            withComment, _
                                            withIfNotExists, _
                                            sSQL, _
                                            sqlCreateFK, _
                                            sqlAddComment)
        Output_WriteLine sSQL & Line
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & Line, OutputID_FirstSection
        End If
        If withComment And Len(sqlAddComment) > 0 Then
            Output_WriteLine sqlAddComment & Line, OutputID_SecondSection
        End If
    Next
End Sub

Public Sub GetSQLDropTableInternal(objLogicalTableCollection As Collection)
    Dim sSQL            As String
    Dim objLogicalTable As clsLogicalTable
    
    '-- Drop tables
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_DropTable_Single(objLogicalTable)
        Output_WriteLine sSQL & Line
    Next
End Sub

Public Sub Get_SQL_CreateTable_Single(objLogicalTable As clsLogicalTable, _
                                ByVal withComment As Boolean, _
                                ByVal withIfNotExists As Boolean, _
                                ByRef sqlCreateTable As String, _
                                ByRef sqlCreateFK As String, _
                                ByRef sqlAddComment As String)
    Dim i               As Integer
    
    Dim syntaxTable         As String
    
    syntaxTable = "CREATE TABLE {6:if not exists}" & IdentifierStart & "{0:table name}" & IdentifierEnd & " (" _
                & "{1:columns definition}" _
                & "{2:primary key}" _
                & "{3:indexes}" _
                & Line & ")" & SQL_Statement_End _
                & "{4:table comment}" _
                & "{5:non unique indexes}"
                
    '-- Create Columns sql
    Dim syntaxColumn        As String
    Dim syntaxDefault       As String
    Dim sqlColumn           As String
    syntaxColumn = "  {0:(i = 1 ? space : ,)}{1:column name} {2:data type} {3:nullable} {4:default}"
    syntaxDefault = "DEFAULT {0:default value}"
    
    sqlColumn = ""
    For i = 1 To objLogicalTable.Columns.Count
        With objLogicalTable.Columns(i)
            sqlColumn = sqlColumn _
                    & Line & FormatString(syntaxColumn, _
                                IIf(i = 1, " ", ","), _
                                .ColumnName, _
                                .DataType, _
                                IIf(.Nullable, "NULL", "NOT NULL"), _
                                FormatString(IIf(Len(.Default) = 0, "", syntaxDefault), StripBrackets(.Default)))
        End With
    Next
    
    '-- Primary key SQL
    Dim syntaxPrimaryKey    As String
    Dim sqlPrimaryKey       As String
    syntaxPrimaryKey = "  ,CONSTRAINT {0:primary key name} PRIMARY KEY ({1:columns})"
    sqlPrimaryKey = ""
    With objLogicalTable.PrimaryKey
        If Len(.PKcolumns) > 0 Then
            sqlPrimaryKey = Line & FormatString(syntaxPrimaryKey, _
                                SQL_Render_PK_Name(objLogicalTable), _
                                 .PKcolumns)
        End If
    End With
    
    '-- Unqiue Indexes SQL
    Dim syntaxUniqueIndex   As String
    Dim sqlUniqueIndex      As String
    syntaxUniqueIndex = "  ,CONSTRAINT {0:index name} UNIQUE {1:columns with bracket}"
    sqlUniqueIndex = ""
    For i = 1 To objLogicalTable.Indexes.Count
        With objLogicalTable.Indexes(i)
            If .IsUnique Then
                sqlUniqueIndex = sqlUniqueIndex _
                        & Line & FormatString(syntaxUniqueIndex, _
                                    SQL_Render_IK_Name(objLogicalTable, objLogicalTable.Indexes(i), True), _
                                    .IKColumns)
            End If
        End With
    Next

    '-- non-unique Indexes SQL
    Dim syntaxNoUniqueIndex As String
    Dim sqlNoUniqueIndex    As String
    syntaxNoUniqueIndex = "CREATE INDEX {0:index name} ON " _
                    & IdentifierStart & "{1:table name}" & IdentifierEnd _
                    & " {2:columns}" & SQL_Statement_End
    sqlNoUniqueIndex = ""
    For i = 1 To objLogicalTable.Indexes.Count
        With objLogicalTable.Indexes(i)
            If Not .IsUnique Then
                sqlNoUniqueIndex = sqlNoUniqueIndex _
                        & Line _
                        & FormatString(syntaxNoUniqueIndex, _
                                SQL_Render_IK_Name(objLogicalTable, objLogicalTable.Indexes(i), False), _
                                objLogicalTable.TableName, _
                                .IKColumns)
            End If
        End With
    Next
    
    '-- table comment
    Dim syntaxTableComment   As String
    Dim sqlTableComment   As String
    syntaxTableComment = "COMMENT ON TABLE {0:table name} IS {1:table comment};"
    sqlTableComment = ""
    If withComment _
        And Len(objLogicalTable.Comment) > 0 Then
        sqlTableComment = Line _
            & FormatString(syntaxTableComment, objLogicalTable.TableName, SQL_ToSQL(objLogicalTable.Comment))
    End If
    
    '-- Generate table sql
    sqlCreateTable = FormatString(syntaxTable, _
                        objLogicalTable.TableName, _
                        sqlColumn, _
                        sqlPrimaryKey, _
                        sqlUniqueIndex, _
                        sqlTableComment, _
                        sqlNoUniqueIndex, _
                        IIf(withIfNotExists, "IF NOT EXISTS ", ""))
    
    '-- Generate Foreign key SQL
    Dim sqlFK           As String
    sqlFK = "ALTER TABLE " & IdentifierStart & "{0:Table Name}" & IdentifierEnd _
            & " ADD CONSTRAINT {1:foreign key name} FOREIGN KEY ({2:column[,..n]}) REFERENCES {3:ref_info}" _
            & SQL_Statement_End
    
    sqlCreateFK = ""
    If objLogicalTable.ForeignKeys.Count > 0 Then
        For i = 1 To objLogicalTable.ForeignKeys.Count
            With objLogicalTable.ForeignKeys(i)
                sqlCreateFK = sqlCreateFK _
                                & IIf(i = 1, "", Line) _
                                & FormatString(sqlFK, _
                                            objLogicalTable.TableName, _
                                            SQL_Render_FK_Name(objLogicalTable, objLogicalTable.ForeignKeys(i)), _
                                            .FKcolumns, _
                                            .RefTableAndColumns & IIf(Len(.fkOption) = 0, "", " " & .fkOption))
            End With
        Next
    End If
    
    '-- Column comment
    Dim syntaxColumnComment   As String
    syntaxColumnComment = "COMMENT ON COLUMN {0:table name}.{1:column name} IS {2:column comment};"
    sqlAddComment = ""
    
    Dim objColumn As clsLogicalColumn
    If withComment Then
        For Each objColumn In objLogicalTable.Columns
            With objColumn
                If Len(.DBDescription) > 0 Then
                    sqlAddComment = sqlAddComment _
                        & Line _
                        & FormatString(syntaxColumnComment, _
                            objLogicalTable.TableName, _
                            .ColumnName, _
                            SQL_ToSQL(.DBDescription))
                End If
            End With
        Next
    End If
    
End Sub

Private Function Get_SQL_DropTable_Single(objLogicalTable As clsLogicalTable) As String
    Dim sSQL            As String
    
    sSQL = "DROP TABLE IF EXISTS " _
            & IdentifierStart & "{0:table name}" & IdentifierEnd _
            & " CASCADE" _
            & SQL_Statement_End
   
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
                                objLogicalIndex As clsLogicalIndex, _
                                IsUnique As Boolean) As String
    SQL_Render_IK_Name = IIf(IsUnique, "UI_", "I_") _
                            & SQL_Render_TableName(objLogicalTable.TableName) _
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
