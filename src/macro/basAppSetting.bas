Attribute VB_Name = "basAppSetting"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2007, 2014, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit
Public Const App_Name                   As String = "Database Modeling Excel"
Public Const App_Version                As String = "7.0"

'-- Sheet part
Public Const Sheet_Index                        As Integer = 1      '-- Index of index sheet
Public Const Sheet_Index_Name                   As String = "~Index~"
Public Const Sheet_Update_History               As Integer = 2      '-- Index of update history sheet
Public Const Sheet_Update_History_Name          As String = "~History~"
Public Const Sheet_Rule                         As Integer = 3      '-- Index of rule sheet
Public Const Sheet_Rule_Name                    As String = "~Rules~"
Public Const Sheet_First_Table                  As Integer = 4      '-- Index of first table sheet

Public Const Table_Sheet_Row_TableName          As Integer = 1
Public Const Table_Sheet_Col_TableName          As String = "B"
Public Const Table_Sheet_Row_TableComment       As Integer = 2
Public Const Table_Sheet_Col_TableComment       As String = "B"
Public Const Table_Sheet_Row_PrimaryKey         As Integer = 3
Public Const Table_Sheet_Col_PrimaryKey         As String = "B"
Public Const Table_Sheet_Row_ForeignKey         As Integer = 4
Public Const Table_Sheet_Col_ForeignKey         As String = "B"
Public Const Table_Sheet_Row_Index              As Integer = 5
Public Const Table_Sheet_Col_Index              As String = "B"
Public Const Table_Sheet_Col_Unique             As String = "G"
Public Const Table_Sheet_Col_Clustered          As String = "H"
Public Const Table_Sheet_Row_TableStatus        As Integer = 3
Public Const Table_Sheet_Col_TableStatus        As String = "I"

Public Const Table_Sheet_Row_First_Column       As Integer = 8
Public Const Table_Sheet_Col_ColumnLabel        As String = "A"
Public Const Table_Sheet_Col_ColumnName         As String = "B"
Public Const Table_Sheet_Col_ColumnDataType     As String = "C"
Public Const Table_Sheet_Col_ColumnNullable     As String = "D"
Public Const Table_Sheet_Col_ColumnDefault      As String = "E"
Public Const Table_Sheet_Col_ColumnComment      As String = "F"

'-- Table Sheet Value
Public Const Table_Sheet_PK_Clustered           As String = ""
Public Const Table_Sheet_PK_NonClustered        As String = "N"
Public Const Table_Sheet_Index_Clustered        As String = "Y"
Public Const Table_Sheet_Index_NonClustered     As String = ""
Public Const Table_Sheet_Index_Unique           As String = ""
Public Const Table_Sheet_Index_NonUnique        As String = "N"
Public Const Table_Sheet_Nullable               As String = "Yes"
Public Const Table_Sheet_NonNullable            As String = "No"
Public Const Table_Sheet_TableStatus_Ignore     As String = "ignore"

'-- UI
Public Const Table_Code_Length                  As Integer = 0
Public Const Sheet_NameIsTableDesc              As Boolean = False

'-- Marks
Public Const Line                               As String = vbCrLf

'-- Databae Type Global variable
Public Const DBName_DB2                         As String = "DB2"
Public Const DBName_MariaDB                     As String = "MariaDB"
Public Const DBName_MySQL                       As String = "MySQL"
Public Const DBName_Oracle                      As String = "Oracle"
Public Const DBName_PostgreSQL                  As String = "PostgreSQL"
Public Const DBName_SQLite                      As String = "SQLite"
Public Const DBName_SQLServer                   As String = "SQL Server"
Public Const DBName_All                         As String = "All"

'----------- Excel Type Global variable ---------------
'-- the constant's value will be one of DBName_All, DBName_SQLServer, DBName_MySQL, or DBName_XXX
Public Const The_Excel_Type                     As String = DBName_All
