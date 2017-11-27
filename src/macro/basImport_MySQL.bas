Attribute VB_Name = "basImport_MySQL"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2012, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

'-------------------------------------------
'-- MySQL Import Module
'-------------------------------------------
Public DRIVER_NAME As String
Public PORT_ID As String
Public SERVER_NAME As String
Public SERVER_DATABASE_NAME As String
Public SERVER_TABLE_NAME As String

Public Function CreateConnection(ByVal Server As String, _
                    ByVal Database As String, _
                    ByVal User As String, _
                    ByVal Password As String, _
                    ByVal driver As String, _
                    ByVal port As String) As ADODB.Connection
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection

    If Len(driver) = 0 Then driver = "{MySQL ODBC 5.2 UNICODE Driver}"
    If Len(port) = 0 Then port = "3306"
    
    conn.connectionString = "Driver=" & Trim(driver) _
            & ";Server=" & Trim(Server) _
            & ";Port=" & Trim(port) _
            & ";Database=" & Database _
            & ";User=" & Trim(User) _
            & ";Password=" & Password _
            & ";OPTION=3"
    
    Set CreateConnection = conn
End Function

Public Function GetLogicalTable(conn As ADODB.Connection, TableName As String) As clsLogicalTable
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
    Set GetLogicalTable = objTable
End Function

Public Sub RenderPKAndIndex(conn As ADODB.Connection, objTable As clsLogicalTable)
    Dim syntax As String
    
    syntax = "   SELECT S.TABLE_NAME" _
    & Line & "        , S.INDEX_NAME" _
    & Line & "        , S.SEQ_IN_INDEX" _
    & Line & "        , S.COLUMN_NAME" _
    & Line & "        , S.NON_UNIQUE" _
    & Line & "        , TC.CONSTRAINT_TYPE" _
    & Line & "     FROM information_schema.STATISTICS S" _
    & Line & "LEFT JOIN information_schema.TABLE_CONSTRAINTS TC" _
    & Line & "       ON S.TABLE_SCHEMA = TC.TABLE_SCHEMA" _
    & Line & "      AND S.TABLE_NAME = TC.TABLE_NAME" _
    & Line & "      AND S.INDEX_NAME = TC.CONSTRAINT_NAME" _
    & Line & "    WHERE S.TABLE_SCHEMA = DATABASE()" _
    & Line & "      AND S.TABLE_NAME = {0:table name}" _
    & Line & " ORDER BY S.TABLE_NAME" _
    & Line & "        , S.INDEX_NAME" _
    & Line & "        , S.SEQ_IN_INDEX;"

    Dim sSQL                    As String
    sSQL = FormatString(syntax, SQL_ToSQL(objTable.TableName))
    
    Dim oRs                     As ADODB.Recordset
    Dim curIndexName            As String
    Dim objIndex                As clsLogicalIndex

    On Error GoTo Flag_Err

    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    curIndexName = ""

    Do While Not oRs.EOF
        If oRs("CONSTRAINT_TYPE") & "" = "PRIMARY KEY" Then
            '-- Primary Key
            If Len(objTable.PrimaryKey.PKcolumns) = 0 Then
                objTable.PrimaryKey.PKcolumns = oRs("COLUMN_NAME") & ""
            Else
                objTable.PrimaryKey.PKcolumns = objTable.PrimaryKey.PKcolumns & ", " & oRs("COLUMN_NAME")
            End If
            objTable.PrimaryKey.IsClustered = True
        Else
            '-- Index
            If curIndexName <> (oRs("INDEX_NAME") & "") Then
                Set objIndex = New clsLogicalIndex
                objTable.Indexes.Add objIndex
                
                objIndex.IsClustered = False
                objIndex.IsUnique = (oRs("NON_UNIQUE") = 0)

                curIndexName = oRs("INDEX_NAME") & ""
            End If

            If Len(objIndex.IKColumns) = 0 Then
                objIndex.IKColumns = oRs("COLUMN_NAME") & ""
            Else
                objIndex.IKColumns = objIndex.IKColumns & ", " & oRs("COLUMN_NAME")
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
    
    syntax = "SELECT R.TABLE_NAME" _
    & Line & "     , R.CONSTRAINT_NAME" _
    & Line & "     , R.UPDATE_RULE" _
    & Line & "     , R.DELETE_RULE" _
    & Line & "     , R.REFERENCED_TABLE_NAME" _
    & Line & "     , K.COLUMN_NAME" _
    & Line & "     , K.ORDINAL_POSITION" _
    & Line & "     , K.POSITION_IN_UNIQUE_CONSTRAINT" _
    & Line & "     , K.REFERENCED_COLUMN_NAME" _
    & Line & "  FROM information_schema.REFERENTIAL_CONSTRAINTS R" _
    & Line & "  JOIN information_schema.KEY_COLUMN_USAGE K" _
    & Line & "    ON R.CONSTRAINT_SCHEMA = K.CONSTRAINT_SCHEMA" _
    & Line & "   AND R.TABLE_NAME        = K.TABLE_NAME" _
    & Line & "   AND R.CONSTRAINT_NAME   = K.CONSTRAINT_NAME" _
    & Line & " WHERE R.CONSTRAINT_SCHEMA = DATABASE()" _
    & Line & "   AND R.TABLE_NAME = {0:table name}" _
    & Line & " ORDER BY R.TABLE_NAME" _
    & Line & "     , R.CONSTRAINT_NAME" _
    & Line & "     , K.ORDINAL_POSITION;"

    Dim sSQL                    As String
    sSQL = FormatString(syntax, SQL_ToSQL(objTable.TableName))
    
    Dim oRs             As ADODB.Recordset
    Dim curFKName       As String
    Dim objForeignKey   As clsLogicalForeignKey
    
    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    curFKName = ""

    Do While Not oRs.EOF
        '-- For Foreign Key
        If curFKName <> (oRs("CONSTRAINT_NAME") & "") Then
            Set objForeignKey = New clsLogicalForeignKey
            objTable.ForeignKeys.Add objForeignKey

            objForeignKey.RefTableName = oRs("REFERENCED_TABLE_NAME")
            If oRs("DELETE_RULE") <> "RESTRICT" Then
                objForeignKey.OnDelete = "ON DELETE " & oRs("DELETE_RULE")
            Else
                objForeignKey.OnDelete = ""
            End If
            If oRs("UPDATE_RULE") <> "RESTRICT" Then
                objForeignKey.OnUpdate = "ON DELETE " & oRs("UPDATE_RULE")
            Else
                objForeignKey.OnUpdate = ""
            End If
            
            curFKName = oRs("CONSTRAINT_NAME") & ""
        End If

        If Len(objForeignKey.FKcolumns) > 0 Then
            objForeignKey.FKcolumns = objForeignKey.FKcolumns & ", "
        End If
        objForeignKey.FKcolumns = objForeignKey.FKcolumns & oRs("COLUMN_NAME")
        
        If Len(objForeignKey.RefTableColumns) > 0 Then
            objForeignKey.RefTableColumns = objForeignKey.RefTableColumns & ", "
        End If
        objForeignKey.RefTableColumns = objForeignKey.RefTableColumns & oRs("REFERENCED_COLUMN_NAME")

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
    
    syntax = "  SELECT C.TABLE_NAME" _
    & Line & "       , C.COLUMN_NAME" _
    & Line & "       , C.ORDINAL_POSITION" _
    & Line & "       , C.COLUMN_TYPE" _
    & Line & "       , C.COLUMN_DEFAULT" _
    & Line & "       , C.EXTRA" _
    & Line & "       , C.IS_NULLABLE" _
    & Line & "       , C.COLUMN_COMMENT" _
    & Line & "       , C.DATA_TYPE" _
    & Line & "       , C.CHARACTER_MAXIMUM_LENGTH" _
    & Line & "       , C.NUMERIC_PRECISION" _
    & Line & "       , C.NUMERIC_SCALE" _
    & Line & "    FROM information_schema.`COLUMNS` C" _
    & Line & "   WHERE C.TABLE_SCHEMA = DATABASE()" _
    & Line & "     AND C.TABLE_NAME = {0:table name}" _
    & Line & "ORDER BY C.TABLE_NAME" _
    & Line & "       , C.ORDINAL_POSITION;"

    Dim sSQL                    As String
    sSQL = FormatString(syntax, SQL_ToSQL(objTable.TableName))
    
    Dim oRs             As ADODB.Recordset
    Dim objColumn       As clsLogicalColumn
    
    '-- Open recordset
    Set oRs = New ADODB.Recordset
    oRs.Open sSQL, conn, adOpenForwardOnly

    Do While Not oRs.EOF
        '-- set Column
        Set objColumn = New clsLogicalColumn
        objTable.Columns.Add objColumn
        
        objColumn.ColumnName = oRs("COLUMN_NAME")
        objColumn.DataType = GetColumnDataType( _
                                            oRs("COLUMN_TYPE"), _
                                            oRs("EXTRA") & "")
        objColumn.Nullable = (oRs("IS_NULLABLE") = "YES")
        objColumn.Default = oRs("COLUMN_DEFAULT") & ""
        objColumn.Comment = oRs("COLUMN_COMMENT") & ""
        
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

Private Function GetColumnDataType(columnType As String, _
                        extra As String) As String
    Dim DataType As String

    DataType = LCase(columnType)
    If (Len(extra) > 0) Then
        DataType = DataType & " " & extra
    End If
    GetColumnDataType = DataType
End Function
