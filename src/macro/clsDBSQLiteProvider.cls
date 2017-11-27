VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBSQLiteProvider"
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
Const IdentifierStart = """"
Const IdentifierEnd = """"
Const OutputID_FirstSection = 1
Const OutputID_SecondSection = 2


'-----------------------------------------------------------
'-- SQLite
'-----------------------------------------------------------

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
                                            sqlAddComment, _
                                            False)
        Output_WriteLine sSQL & Line
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & Line & SQL_Statement_End, OutputID_FirstSection
        End If
        If withComment Then
            If Len(sqlAddComment) > 0 Then
                Output_WriteLine sqlAddComment & Line & SQL_Statement_End, OutputID_SecondSection
            End If
        End If
    Next
    
    Output_Copy
End Sub

Public Sub clsIDatabaseProvider_GetSQLDropTable(objLogicalTableCollection As Collection, _
                            Optional outputOptions As clsOutputOptions)
    Dim sSQL            As String
    Dim objLogicalTable As clsLogicalTable
    
    Output_Initialize outputOptions
    
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
    Dim sqlAddComment       As String
    Dim objLogicalTable     As clsLogicalTable
    
    Output_Initialize outputOptions
    
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
                                            sqlAddComment, _
                                            False)
        Output_WriteLine sSQL & Line
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & Line & SQL_Statement_End, 1
        End If
        If withComment Then
            If Len(sqlAddComment) > 0 Then
                Output_WriteLine sqlAddComment & Line & SQL_Statement_End, 2
            End If
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
                                            sSQL)
        Output_WriteLine sSQL & Line
    Next
    
    Output_Copy
End Sub

Private Sub Get_SQL_CreateTableIfNotExists_Single(objLogicalTable As clsLogicalTable, _
                                ByRef sqlCreateTable As String)
    Call Get_SQL_CreateTable_Single(objLogicalTable, _
                                False, _
                                sqlCreateTable, _
                                "", _
                                "", _
                                True)
End Sub

Public Sub Get_SQL_CreateTable_Single(objLogicalTable As clsLogicalTable, _
                                ByVal withComment As Boolean, _
                                ByRef sqlCreateTable As String, _
                                ByRef sqlCreateFK As String, _
                                ByRef sqlAddComment As String, _
                                ByVal IfNotExists As Boolean)
    Dim i                   As Integer
    Dim syntaxTable         As String
    
    syntaxTable = "CREATE TABLE {6:if not exists}" & IdentifierStart & "{0:table name}" & IdentifierEnd & " (" _
                & "{1:columns definition}" _
                & "{2:primary key}" _
                & "{3:unique indexes}" _
                & "{4:foreign keys}" _
                & Line & ")" & SQL_Statement_End _
                & "{5:non unique indexes}"
    
    '-- Create Columns sql
    Dim syntaxColumn        As String
    Dim syntaxDefault       As String
    Dim sqlColumn           As String
        
    syntaxColumn = "  {0:(i = 1 ? space : ,)}" & IdentifierStart & "{1:column name}" & IdentifierEnd & " {2:data type} {3:nullable} {4:default}"
    syntaxDefault = "  CONSTRAINT {0:default name} DEFAULT ({1:default value})"

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
                                    .Default))
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
    
    syntaxUniqueIndex = "  ,CONSTRAINT {0:index name} UNIQUE {1:columns with brackets}"
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
    Dim syntaxNoUniqueIndex As String
    Dim sqlNoUniqueIndex    As String
    
    syntaxNoUniqueIndex = "CREATE INDEX {3:if not exists}{0:index name} ON " _
                    & IdentifierStart & "{1:table name}" & IdentifierEnd _
                    & " {2:columns with brackets}" & SQL_Statement_End
    sqlNoUniqueIndex = ""
    For i = 1 To objLogicalTable.Indexes.Count
        With objLogicalTable.Indexes(i)
            If Not .IsUnique Then
                sqlNoUniqueIndex = sqlNoUniqueIndex _
                        & Line & FormatString(syntaxNoUniqueIndex, _
                                    SQL_Render_IK_Name(objLogicalTable, objLogicalTable.Indexes(i)), _
                                    objLogicalTable.TableName, _
                                    .IKColumns, _
                                    IIf(IfNotExists, "IF NOT EXISTS ", ""))
            End If
        End With
    Next
    
    '-- Generate Foreign key SQL
    sqlCreateFK = ""    '-- SQLite does not support Add FK in Alter Table statement.
    Dim syntaxForeignKey    As String
    Dim sqlForeignKey       As String
    
    syntaxForeignKey = "  ,CONSTRAINT {0:foreign key name} FOREIGN KEY ({1:column[,..n]}) REFERENCES {2:ref_info}"
    sqlForeignKey = ""
    For i = 1 To objLogicalTable.ForeignKeys.Count
        With objLogicalTable.ForeignKeys(i)
            sqlForeignKey = sqlForeignKey _
                            & Line & FormatString(syntaxForeignKey, _
                                        SQL_Render_FK_Name(objLogicalTable, objLogicalTable.ForeignKeys(i)), _
                                        .FKcolumns, _
                                        .RefTableAndColumns & IIf(Len(.fkOption) = 0, "", " " & .fkOption))
        End With
    Next
    
    '-- Generate table sql
    sqlCreateTable = FormatString(syntaxTable, _
                        objLogicalTable.TableName, _
                        sqlColumn, _
                        sqlPrimaryKey, _
                        sqlUniqueIndex, _
                        sqlForeignKey, _
                        sqlNoUniqueIndex, _
                        IIf(IfNotExists, "IF NOT EXISTS ", ""))
    
'    YN: NOT Support
'   '-- Column comment
   
End Sub

Private Function Get_SQL_DropTable_Single(objLogicalTable As clsLogicalTable) As String
    Dim sSQL            As String
    
    sSQL = "DROP TABLE IF EXISTS " _
        & IdentifierStart & "{0:table name}" & IdentifierEnd _
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
