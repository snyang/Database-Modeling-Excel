VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBOracleProvider"
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
'-- Oracle
'-----------------------------------------------------------
Private Const SP_REMVOE_FK_NAME = "tmp_dbmodelexcel_drop_table_fk"
Private Const MAX_OBJECT_NAME_LENGTH = 30
Private Const SQL_DELIMITER = "/"
Private p_colObjectNames    As Collection
'-------------------------------------

Public Sub clsIDatabaseProvider_GetSQLCreateTable(objLogicalTableCollection As Collection, _
                            ByVal withComment As Boolean, _
                            Optional outputOptions As clsOutputOptions)
    Dim sSQL                As String
    Dim sqlCreateFK         As String
    Dim sqlAddComment       As String
    Dim objLogicalTable     As clsLogicalTable
    
    Set p_colObjectNames = New Collection
    Output_Initialize outputOptions
    
    '-- Create tables
    For Each objLogicalTable In objLogicalTableCollection
        Call Get_SQL_CreateTable_Single(objLogicalTable, _
                                            withComment, _
                                            sSQL, _
                                            sqlCreateFK, _
                                            sqlAddComment)
        Output_WriteLine sSQL
        If withComment Then
            Output_WriteLine sqlAddComment
        End If
        Output_WriteLine ""
        
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & Line, 1
        End If
    Next
    
    Output_Copy
    Set p_colObjectNames = Nothing
End Sub

Public Sub clsIDatabaseProvider_GetSQLDropTable(objLogicalTableCollection As Collection, _
                            Optional outputOptions As clsOutputOptions)
    Dim sSQL            As String
    Dim objLogicalTable As clsLogicalTable
    
    Output_Initialize outputOptions
    
    '-- Create the store procedue of dropping foreign key
    sSQL = Get_SQL_Pre_RemoveFK()
    Output_WriteLine sSQL & Line & SQL_DELIMITER
    
    '-- Drop foreign key relation
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_RemoveFK_Single(objLogicalTable)
        Output_WriteLine sSQL & Line
    Next
    
    '-- Drop the store procedue of dropping foreign key
    sSQL = Get_SQL_Post_RemoveFK()
    Output_WriteLine sSQL & Line
    
    '-- Drop tables
    Output_WriteLine Get_SQL_DropTable_Single_Pre
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_DropTable_Single(objLogicalTable)
        Output_WriteLine sSQL & Line
    Next
    Output_WriteLine Get_SQL_DropTable_Single_Post
    
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
    
    '-- Create the store procedue of dropping foreign key
    sSQL = Get_SQL_Pre_RemoveFK()
    Output_WriteLine sSQL & Line & SQL_DELIMITER
    
    '-- Drop foreign key relation
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_RemoveFK_Single(objLogicalTable)
        Output_WriteLine sSQL & Line
    Next
    
    '-- Drop the store procedue of dropping foreign key
    sSQL = Get_SQL_Post_RemoveFK()
    Output_WriteLine sSQL & Line
    
    '-- Drop tables
    Output_WriteLine Get_SQL_DropTable_Single_Pre
    For Each objLogicalTable In objLogicalTableCollection
        sSQL = Get_SQL_DropTable_Single(objLogicalTable)
        Output_WriteLine sSQL & Line
    Next
    Output_WriteLine Get_SQL_DropTable_Single_Post
    
    '-- Create tables
    Set p_colObjectNames = New Collection
    For Each objLogicalTable In objLogicalTableCollection
        Call Get_SQL_CreateTable_Single(objLogicalTable, _
                                            withComment, _
                                            sSQL, _
                                            sqlCreateFK, _
                                            sqlAddComment)
        Output_WriteLine sSQL
        If withComment Then
            Output_WriteLine sqlAddComment
        End If
        Output_WriteLine ""
        
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & Line, 1
        End If
    Next
    
    Output_Copy
    Set p_colObjectNames = Nothing
End Sub

Public Sub clsIDatabaseProvider_GetSQLCreateTableIfNotExists(objLogicalTableCollection As Collection, _
                            Optional outputOptions As clsOutputOptions)
    Dim sSQL                As String
    Dim sqlCreateFK         As String
    Dim objLogicalTable     As clsLogicalTable
    
    Set p_colObjectNames = New Collection
    Output_Initialize outputOptions
    
    Output_WriteLine Get_SQL_CreateTableIfNotExists_Single_Pre
    
    For Each objLogicalTable In objLogicalTableCollection
        Call Get_SQL_CreateTableIfNotExists_Single(objLogicalTable, _
                                            sSQL, _
                                            sqlCreateFK)
        Output_WriteLine sSQL & Line
        If Len(sqlCreateFK) > 0 Then
            Output_WriteLine sqlCreateFK & Line, 1
        End If
    Next
    
    Output_WriteLine Get_SQL_CreateTableIfNotExists_Single_Post
    
    Output_Copy
    Set p_colObjectNames = Nothing
End Sub

Private Sub Get_SQL_CreateTableIfNotExists_Single(objLogicalTable As clsLogicalTable, _
                                ByRef sqlCreateTable As String, _
                                ByRef sqlCreateFK As String)
    Dim sSQL            As String
    
    sSQL = "  SELECT COUNT(*) INTO v_table_is_exists" _
        & Line & "  FROM user_tables" _
        & Line & "  WHERE lower(table_name) = lower('{0:table name}');" _
        & Line & "  IF v_table_is_exists != 0 THEN" _
        & Line & "    execute immediate {2:create table sql}" _
        & Line & "  END IF;"
    
    Call Get_SQL_CreateTable_Single(objLogicalTable, _
                                False, _
                                sqlCreateTable, _
                                sqlCreateFK, _
                                "")
                                
    sqlCreateTable = FormatString(sSQL, _
                    objLogicalTable.TableName, _
                    objLogicalTable.Comment, _
                    Replace(sqlCreateTable, "'", "''"))
    
End Sub

Private Function Get_SQL_CreateTableIfNotExists_Single_Pre() As String
    Dim sSQL            As String
    
    sSQL = "DECLARE" _
        & Line & "  v_table_is_exists integer;" _
        & Line & "BEGIN" _

    '-- Return
    Get_SQL_CreateTableIfNotExists_Single_Pre = sSQL
End Function

Private Function Get_SQL_CreateTableIfNotExists_Single_Post() As String
    Dim sSQL            As String
    
    sSQL = "END;" & Line & SQL_DELIMITER
    
    '-- Return
    Get_SQL_CreateTableIfNotExists_Single_Post = sSQL
End Function

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
            & Line & "CREATE or REPLACE PROCEDURE {0:sp name for remove FK}(" _
            & Line & "    a_table_name IN VARCHAR2" _
            & Line & ") IS" _
            & Line & "  v_fk_name varchar2(250);" _
            & Line & "  CURSOR c_fk IS" _
            & Line & "    SELECT UC.constraint_name" _
            & Line & "      FROM user_constraints UC" _
            & Line & "     WHERE lower(UC.table_name) = lower(a_table_name)" _
            & Line & "       AND UC.constraint_type = 'R';" _
            & Line & "BEGIN" _
            & Line & "" _
            & Line & "  OPEN c_fk;" _
            & Line & "  LOOP" _
            & Line & "    FETCH c_fk INTO v_fk_name;" _
            & Line & "    EXIT WHEN c_fk%NOTFOUND;" _
            & Line & "      EXECUTE IMMEDIATE 'ALTER TABLE ' || a_table_name || ' DROP CONSTRAINT ' || v_fk_name;" _
            & Line & "    END LOOP;" _
            & Line & "  CLOSE c_fk;" _
            & Line & "END {0:sp name for remove FK};"
    '-- Return
    Get_SQL_Pre_RemoveFK = FormatString(sSQL, SP_REMVOE_FK_NAME)
End Function

Private Function Get_SQL_Post_RemoveFK() As String
    Dim sSQL        As String
    
    sSQL = "-- Remove temporary store procedue for remove foreign key" _
            & Line & "DROP PROCEDURE {0:sp name for remove FK};" _
            
    '-- Return
    Get_SQL_Post_RemoveFK = FormatString(sSQL, SP_REMVOE_FK_NAME)
End Function

Public Sub Get_SQL_CreateTable_Single(objLogicalTable As clsLogicalTable, _
                                ByVal withComment As Boolean, _
                                ByRef sqlCreateTable As String, _
                                ByRef sqlCreateFK As String, _
                                ByRef sqlAddComment As String)
    Dim i               As Integer
    
    Dim syntaxTable         As String
    Dim syntaxColumn        As String
    Dim syntaxDefault       As String
    Dim syntaxPrimaryKey    As String
    Dim syntaxUniqueIndex   As String
    Dim syntaxNoUniqueIndex As String
    '-- Create table
    'create table TEST
    '(
    '  ID   NUMBER not null,
    '  NAME VARCHAR2(20) default 'y'
    ')
    ';
    
    syntaxTable = VBA.String(32, "-") _
                & Line & "-- Create table: '{0:table name}'" _
                & Line & VBA.String(32, "-") _
                & Line & "CREATE TABLE {0:table name} (" _
                & "{1:columns definition}" _
                & Line & ");" _
                & "{2: primary key}" _
                & "{3: index}" _
                & "{4: non-unique index}"
    syntaxColumn = "  {0:(i = 1 ? space : ,)}{1:column name} {2:data type}{4:default}{3:nullable}"
    syntaxDefault = " DEFAULT {0:default value}"
    
    syntaxPrimaryKey = "ALTER TABLE {0:table name} ADD CONSTRAINT {1:primary key name} PRIMARY KEY ({2:columns});"
    syntaxUniqueIndex = "ALTER TABLE {0:table name} ADD CONSTRAINT {1:unique index name} UNIQUE {2:columns with bracket};"
    syntaxNoUniqueIndex = "CREATE INDEX {1:index name} ON {0:table name} {2:columns};"
    
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
                                IIf(.Nullable, " NULL", " NOT NULL"), _
                                FormatString(IIf(Len(.Default) = 0, "", syntaxDefault), _
                                    .Default))
        End With
    Next
    
    '-- Primary key SQL
    sqlPrimaryKey = ""
    With objLogicalTable.PrimaryKey
        If Len(.PKcolumns) > 0 Then
            sqlPrimaryKey = Line & FormatString(syntaxPrimaryKey, _
                                objLogicalTable.TableName, _
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
                                    objLogicalTable.TableName, _
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
                        & Line & FormatString(syntaxNoUniqueIndex, _
                                    objLogicalTable.TableName, _
                                    SQL_Render_IK_Name(objLogicalTable, objLogicalTable.Indexes(i)), _
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
    sqlCreateFK = ""
    Dim sqlFK           As String
    sqlFK = "ALTER TABLE {0:Table Name}" _
        & Line & "  ADD CONSTRAINT {1:foreign key name}" _
        & Line & "  FOREIGN KEY ({2:column[,..n]})" _
        & Line & "  REFERENCES {3:ref_info};"
    
    For i = 1 To objLogicalTable.ForeignKeys.Count
        With objLogicalTable.ForeignKeys(i)
            sqlCreateFK = sqlCreateFK _
                    & Line & FormatString(sqlFK, _
                                objLogicalTable.TableName, _
                                SQL_Render_FK_Name(objLogicalTable, objLogicalTable.ForeignKeys(i)), _
                                .FKcolumns, _
                                .RefTableAndColumns & IIf(Len(.fkOption) = 0, "", " " & .fkOption))
        End With
    Next
    
    If Len(sqlCreateFK) > 0 Then
        sqlCreateFK = FormatString("-- Create foreign keys for table: '{0:table name}'", objLogicalTable.TableName) _
             & sqlCreateFK
    End If
    
    '-- Column comment
    Dim syntaxTableComment          As String
    syntaxTableComment = "COMMENT ON TABLE {0:table name} IS {1:table comment};"
    Dim syntaxColumnComment   As String
    syntaxColumnComment = "COMMENT ON COLUMN {0:table name}.{1:column name} IS {2:column comment};"
    
    sqlAddComment = FormatString(syntaxTableComment, _
                                objLogicalTable.TableName, _
                                SQL_ToSQL(objLogicalTable.Comment))
    
    Dim strComment As String
    Dim objColumn As clsLogicalColumn
    If withComment Then
        For i = 1 To objLogicalTable.Columns.Count
            Set objColumn = objLogicalTable.Columns(i)
            With objColumn
                If Len(.DBDescription) > 0 Then
                    sqlAddComment = sqlAddComment _
                        & Line & FormatString(syntaxColumnComment, _
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
    
    sSQL = "  SELECT COUNT(*) INTO v_table_is_exists" _
        & Line & "  FROM user_tables" _
        & Line & "  WHERE lower(table_name) = lower('{0:table name}');" _
        & Line & "  IF v_table_is_exists != 0 THEN" _
        & Line & "    execute immediate 'DROP TABLE {0:table name} PURGE';" _
        & Line & "  END IF;" _
   
    '-- Return
    Get_SQL_DropTable_Single = FormatString(sSQL, _
                                objLogicalTable.TableName, _
                                objLogicalTable.Comment)
End Function

Private Function Get_SQL_DropTable_Single_Pre() As String
    Dim sSQL            As String
    
    sSQL = "DECLARE" _
        & Line & "  v_table_is_exists integer;" _
        & Line & "BEGIN"

    '-- Return
    Get_SQL_DropTable_Single_Pre = sSQL
End Function

Private Function Get_SQL_DropTable_Single_Post() As String
    Dim sSQL            As String
    
    sSQL = "END;" & Line & "/" & Line
    
    '-- Return
    Get_SQL_DropTable_Single_Post = sSQL
End Function

Private Function SQL_Render_TableName(ByVal TableName As String) As String
    SQL_Render_TableName = Replace(Replace(TableName, " ", ""), "_", "")
End Function

Private Function SQL_Render_PK_Name(objLogicalTable As clsLogicalTable) As String
    SQL_Render_PK_Name = RenderObjectName( _
                            "PK_" & SQL_Render_TableName(objLogicalTable.TableName) _
                            )
End Function

Private Function SQL_Render_FK_Name(objLogicalTable As clsLogicalTable, _
                                objLogicalForeignKey As clsLogicalForeignKey) As String
    SQL_Render_FK_Name = RenderObjectName("FK_" & SQL_Render_TableName(objLogicalTable.TableName) _
                            & "_" & Replace(Replace(Replace(objLogicalForeignKey.FKcolumns, " ", ""), "_", ""), ",", "_") _
                            )
End Function

Private Function SQL_Render_IK_Name(objLogicalTable As clsLogicalTable, _
                                objLogicalIndex As clsLogicalIndex) As String
    SQL_Render_IK_Name = RenderObjectName("IK_" & SQL_Render_TableName(objLogicalTable.TableName) _
                            & "_" & Replace(Replace(Replace(Replace(Replace(objLogicalIndex.IKColumns, _
                                                                    " ", ""), _
                                                            "(", ""), _
                                                    ")", ""), _
                                            "_", ""), _
                                    ",", "_") _
                        )
                                    
End Function

Private Function SQL_Render_DF_Name(objLogicalTable As clsLogicalTable, _
                                objLogicalColumn As clsLogicalColumn) As String
    SQL_Render_DF_Name = RenderObjectName( _
                    "DF_" & SQL_Render_TableName(objLogicalTable.TableName) & "_" & objLogicalColumn.ColumnName _
                    )
                    
End Function

Private Function RenderObjectName(ByVal ObjectName As String) As String
    Dim newObjectName           As String
    Dim objectNameIndex         As Integer
    newObjectName = ObjectName
    
    If Len(newObjectName) > MAX_OBJECT_NAME_LENGTH Then
        newObjectName = Left(newObjectName, MAX_OBJECT_NAME_LENGTH - 4)
        On Error Resume Next
        objectNameIndex = p_colObjectNames.Item(newObjectName)
        If Err.Number <> 0 Then
            Err.Clear
            objectNameIndex = 0
            Call p_colObjectNames.Add(objectNameIndex, newObjectName)
        Else
            objectNameIndex = objectNameIndex + 1
            p_colObjectNames.Item(newObjectName) = objectNameIndex
        End If
        newObjectName = newObjectName & Format(objectNameIndex, "0000")
    End If
    
    '-- Return
    RenderObjectName = newObjectName
End Function



