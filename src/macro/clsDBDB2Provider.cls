VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBDB2Provider"
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
'-- DB2
'Syntax:
'  Comment: -- or /* */
'  String literal modifier: '
'  Identifier modifier: <not sure>
'-----------------------------------------------------------

Const SQL_Statement_End = " @"
Const OutputID_FirstSection = 1

Public Sub clsIDatabaseProvider_GetSQLCreateTable(objLogicalTableCollection As Collection, _
                            ByVal withComment As Boolean, _
                            Optional outputOptions As clsOutputOptions)
    
    Output_Initialize outputOptions
    
    OutputSetTerminator
    
    Call GetSQLCreateTableInternal(objLogicalTableCollection, withComment, False)
    
    Output_Copy
End Sub

Public Sub clsIDatabaseProvider_GetSQLDropTable(objLogicalTableCollection As Collection, _
                            Optional outputOptions As clsOutputOptions)
    Output_Initialize outputOptions
    
    OutputSetTerminator
    
    '-- Drop table SQL
    Call GetSQLDropTableInternal(objLogicalTableCollection)

    Output_Copy
End Sub

Public Sub clsIDatabaseProvider_GetSQLDropAndCreateTable(objLogicalTableCollection As Collection, _
                            ByVal withComment As Boolean, _
                            Optional outputOptions As clsOutputOptions)
    Output_Initialize outputOptions
    
    OutputSetTerminator
    
    '-- Drop table SQL
    Call GetSQLDropTableInternal(objLogicalTableCollection)
    
    '-- Create tables SQL
    Call GetSQLCreateTableInternal(objLogicalTableCollection, withComment, False)
    
    Output_Copy
End Sub

Public Sub clsIDatabaseProvider_GetSQLCreateTableIfNotExists(objLogicalTableCollection As Collection, _
                            Optional outputOptions As clsOutputOptions)
    Output_Initialize outputOptions
    
    OutputSetTerminator
    
    Call GetSQLCreateTableInternal(objLogicalTableCollection, False, True)
    
    Output_Copy
End Sub

Private Sub OutputSetTerminator()
    Dim sSQL As String
    sSQL = "--#SET TERMINATOR" & SQL_Statement_End
    Output_WriteLine sSQL & Line
End Sub

Private Sub GetSQLCreateTableInternal(objLogicalTableCollection As Collection, _
                            ByVal withComment As Boolean, _
                            ByVal withIfNotExists As Boolean)
    
    Dim sSQL                As String
    Dim sqlCreateFK         As String
    Dim objLogicalTable     As clsLogicalTable
    '-- Create tables
    For Each objLogicalTable In objLogicalTableCollection
        Call Get_SQL_CreateTable_Single(objLogicalTable, _
                                            withComment, _
                                            withIfNotExists, _
                                            sSQL, _
                                            sqlCreateFK)
        Output_WriteLine sSQL & Line
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & Line, OutputID_FirstSection
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
                                ByRef sqlCreateFK As String)
    
    Dim objColumn           As clsLogicalColumn
    Dim objIndex            As clsLogicalIndex
    Dim objForeignKey       As clsLogicalForeignKey
    Dim isFirst             As Boolean
    Dim syntaxTable         As String
    
    syntaxTable = "-- Create table '{0:table name}'" _
         & Line & "CREATE TABLE {5:if not exists}" & "{0:table name}" & " (" _
                & "{1:columns definition}" _
                & "{2:primary key}" _
         & Line & ")" & SQL_Statement_End _
                & "{3:indexes}" _
                & "{4:non unique indexes}"
                
    '-- Create Columns sql
    Dim syntaxColumn        As String
    Dim syntaxDefault       As String
    Dim sqlColumn           As String
    syntaxColumn = "  {0:(i = 1 ? space : ,)}{1:column name} {2:data type} {3:nullable} {4:default}"
    syntaxDefault = "DEFAULT {0:default value}"
    
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
                                FormatString(IIf(Len(.Default) = 0, "", syntaxDefault), .Default))
        End With
        isFirst = False
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
    syntaxUniqueIndex = "CREATE UNIQUE INDEX {0:index name} ON {1:table name} {2:columns}{3:cluster}" & SQL_Statement_End
    sqlUniqueIndex = ""
    For Each objIndex In objLogicalTable.Indexes
        With objIndex
            If .IsUnique Then
                sqlUniqueIndex = sqlUniqueIndex _
                    & Line _
                    & FormatString(syntaxUniqueIndex, _
                        SQL_Render_IK_Name(objLogicalTable, objIndex, True), _
                        objLogicalTable.TableName, _
                        .IKColumns, _
                        IIf(.IsClustered, " CLUSTER", ""))
            End If
        End With
    Next

    '-- non-unique Indexes SQL
    Dim syntaxNoUniqueIndex As String
    Dim sqlNoUniqueIndex    As String
    syntaxNoUniqueIndex = "CREATE INDEX {0:index name} ON {1:table name} {2:columns}{3:cluster}" & SQL_Statement_End
    sqlNoUniqueIndex = ""
    For Each objIndex In objLogicalTable.Indexes
        With objIndex
            If Not .IsUnique Then
                sqlNoUniqueIndex = sqlNoUniqueIndex _
                    & Line _
                    & FormatString(syntaxNoUniqueIndex, _
                        SQL_Render_IK_Name(objLogicalTable, objIndex, False), _
                        objLogicalTable.TableName, _
                        .IKColumns, _
                        IIf(.IsClustered, "CLUSTER", ""))
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
                        IIf(withIfNotExists, "", ""))
    
    '-- Generate Foreign key SQL
    sqlCreateFK = ""
    
    Dim sqlFKHead       As String
    sqlFKHead = "-- Foreign keys for table '{0:table name}'" & Line
    If objLogicalTable.ForeignKeys.Count > 0 Then
        sqlCreateFK = FormatString(sqlFKHead, objLogicalTable.TableName)
    End If
    
    Dim sqlFK           As String
    sqlFK = "ALTER TABLE " & "{0:Table Name}" _
            & Line & "  ADD CONSTRAINT {1:foreign key name}" _
            & Line & "  FOREIGN KEY ({2:column[,..n]})" _
            & Line & "  REFERENCES {3:ref_info}" _
            & SQL_Statement_End
    isFirst = True
    For Each objForeignKey In objLogicalTable.ForeignKeys
        With objForeignKey
            sqlCreateFK = sqlCreateFK _
                            & IIf(isFirst, "", Line) _
                            & FormatString(sqlFK, _
                                        objLogicalTable.TableName, _
                                        SQL_Render_FK_Name(objLogicalTable, objForeignKey), _
                                        .FKcolumns, _
                                        .RefTableAndColumns & IIf(Len(.fkOption) = 0, "", " " & .fkOption))
        End With
        isFirst = False
    Next
    
    '-- Table comment
    Dim sqlComment          As String
    sqlComment = ""
    
    Dim syntaxTableComment  As String
    syntaxTableComment = "COMMENT ON TABLE {0:table name} IS {1:table comment}" & SQL_Statement_End
    If withComment _
        And Len(objLogicalTable.Comment) > 0 Then
        sqlComment = Line _
            & FormatString(syntaxTableComment, objLogicalTable.TableName, SQL_ToSQL(objLogicalTable.Comment))
    End If
    
    '-- Column comment
    Dim syntaxColumnComment   As String
    syntaxColumnComment = "COMMENT ON COLUMN {0:table name}.{1:column name} IS {2:column comment}" & SQL_Statement_End
    
    If withComment Then
        For Each objColumn In objLogicalTable.Columns
            With objColumn
                If Len(.DBDescription) > 0 Then
                    sqlComment = sqlComment _
                        & Line _
                        & FormatString(syntaxColumnComment, _
                            objLogicalTable.TableName, _
                            .ColumnName, _
                            SQL_ToSQL(.DBDescription))
                End If
            End With
        Next
    End If
    
    If Len(sqlComment) > 0 Then
        sqlCreateTable = sqlCreateTable _
                        & Line _
                        & sqlComment
    End If
    
End Sub

Private Function Get_SQL_DropTable_Single(objLogicalTable As clsLogicalTable) As String
    Dim syntax            As String
    
    syntax = "-- Drop table '{0:table name}'" _
    & Line & "BEGIN" _
    & Line & "  IF EXISTS (SELECT * FROM SYSCAT.TABLES" _
    & Line & "             WHERE TABSCHEMA = CURRENT_SCHEMA" _
    & Line & "             AND TABNAME = '{0: table name in upper case}') THEN" _
    & Line & "    EXECUTE IMMEDIATE 'DROP TABLE {0:table name}';" _
    & Line & "  END IF;" _
    & Line & "END" & SQL_Statement_End

   
    '-- Return
    Get_SQL_DropTable_Single = FormatString(syntax, UCase(objLogicalTable.TableName))
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


