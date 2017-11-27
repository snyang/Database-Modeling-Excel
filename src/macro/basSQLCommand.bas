Attribute VB_Name = "basSQLCommand"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2012, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

'-- Copy create all tables SQL to clipboard
Public Sub CopyAllCreateTableSQL()
    Call basPublicDatabase.GetDatabaseProvider().GetSQLCreateTable(GetAllLogicalTables(), False)
End Sub

'-- Copy create all tables with comment SQL to clipboard
Public Sub CopyAllCreateTableWithDescriptionSQL()
    Call basPublicDatabase.GetDatabaseProvider().GetSQLCreateTable(GetAllLogicalTables(), True)
End Sub

'-- Copy drop all tables SQL to clipboard
Public Sub CopyAllDropTableSQL()
    Call basPublicDatabase.GetDatabaseProvider().GetSQLDropTable(GetAllLogicalTables())
End Sub

'-- Copy drop and create all tables SQL to clipboard
Public Sub CopyAllDropAndCreateTableSQL()
    Call basPublicDatabase.GetDatabaseProvider().GetSQLDropAndCreateTable(GetAllLogicalTables(), False)
End Sub

'-- Copy create all exits tables SQL to clipboard
Public Sub CopyAllCreateTableIfNotExistsSQL()
    Call basPublicDatabase.GetDatabaseProvider().GetSQLCreateTableIfNotExists(GetAllLogicalTables())
End Sub

'-- Copy drop and create all tables with comment SQL to clipboard
Public Sub CopyAllDropAndCreateTableWithDescriptionSQL()
    Call basPublicDatabase.GetDatabaseProvider().GetSQLDropAndCreateTable(GetAllLogicalTables(), True)
End Sub

'-- Save create all tables SQL
Public Sub SaveAllCreateTableSQL(fileName As String)
    Call basPublicDatabase.GetDatabaseProvider().GetSQLCreateTable(GetAllLogicalTables(), False, _
            CreateOutputOptions(opmToFile, fileName))
End Sub

'-- Save create all tables with comment SQL
Public Sub SaveAllCreateTableWithDescriptionSQL(fileName As String)
    Call basPublicDatabase.GetDatabaseProvider().GetSQLCreateTable(GetAllLogicalTables(), True, _
            CreateOutputOptions(opmToFile, fileName))
End Sub

'-- Save drop all tables SQL
Public Sub SaveAllDropTableSQL(fileName As String)
    Call basPublicDatabase.GetDatabaseProvider().GetSQLDropTable(GetAllLogicalTables(), _
            CreateOutputOptions(opmToFile, fileName))
End Sub

'-- Save drop and create all tables SQL
Public Sub SaveAllDropAndCreateTableSQL(fileName As String)
    Call basPublicDatabase.GetDatabaseProvider().GetSQLDropAndCreateTable(GetAllLogicalTables(), False, _
            CreateOutputOptions(opmToFile, fileName))
End Sub

'-- Save create all exits tables SQL
Public Sub SaveAllCreateTableIfNotExistsSQL(fileName As String)
    Call basPublicDatabase.GetDatabaseProvider().GetSQLCreateTableIfNotExists(GetAllLogicalTables(), _
            CreateOutputOptions(opmToFile, fileName))
End Sub

'-- Save script of drop and create all tables with comment SQL
Public Sub SaveAllDropAndCreateTableWithDescriptionSQL(fileName As String)
    Call basPublicDatabase.GetDatabaseProvider().GetSQLDropAndCreateTable(GetAllLogicalTables(), _
            True, _
            CreateOutputOptions(opmToFile, fileName))
End Sub
