Attribute VB_Name = "basPublicDatabase"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2014, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

'---------------------------------------------------------------------
'-- Cet connection string form the OLE DB Links dialog.
'---------------------------------------------------------------------
Public Function GetConnectionString(connString As String)
    Dim conn As ADODB.Connection
    Dim MSDASCObj As MSDASC.DataLinks
    
    Set MSDASCObj = New MSDASC.DataLinks
    
    Set conn = New ADODB.Connection
    conn.connectionString = connString
    
    On Error Resume Next
    MSDASCObj.PromptEdit conn
    If Err.Number <> 0 Then
        MSDASCObj.PromptNew conn
    End If
  
    On Error GoTo 0
    GetConnectionString = conn.connectionString
End Function

Public Function GetDatabaseProvider(Optional DatabaseType As String) As clsIDatabaseProvider
    If IsMissing(DatabaseType) Or Len(DatabaseType) = 0 Then
        DatabaseType = The_Excel_Type
    End If
    
    Select Case DatabaseType
    Case DBName_All
        Set GetDatabaseProvider = New clsDBSQLServerProvider
    Case DBName_DB2
        Set GetDatabaseProvider = New clsDBDB2Provider
    Case DBName_MariaDB
        Set GetDatabaseProvider = New clsDBMariaDBProvider
    Case DBName_MySQL
        Set GetDatabaseProvider = New clsDBMySQLProvider
    Case DBName_Oracle
        Set GetDatabaseProvider = New clsDBOracleProvider
    Case DBName_PostgreSQL
        Set GetDatabaseProvider = New clsDBPostgreSQLProvider
    Case DBName_SQLite
        Set GetDatabaseProvider = New clsDBSQLiteProvider
    Case DBName_SQLServer
        Set GetDatabaseProvider = New clsDBSQLServerProvider
    End Select
End Function

Public Function GetImportProvider(DatabaseType As String) As IImportProvider
   
    Select Case DatabaseType
    Case DBName_DB2
        Set GetImportProvider = New clsImportProviderDB2
    Case DBName_MariaDB
        Set GetImportProvider = New clsImportProviderMariaDB
    Case DBName_Oracle
        Set GetImportProvider = New clsImportProviderOracle
    Case DBName_PostgreSQL
        Set GetImportProvider = New clsImportProviderPostgreSQL
    Case DBName_SQLServer
        Set GetImportProvider = New clsImportProviderSQLServer
    End Select
End Function

