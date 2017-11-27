VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IImportProvider"
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
'-------------------------------------------
'-- Oracle Import Provider
'-------------------------------------------
Option Explicit

Public Property Get providers() As String()
End Property

Public Property Get SupportSelectDatabase() As Boolean
End Property

Public Function GetDatabasesSql() As String
End Function

Public Function GetDatabases(ByVal Server As String, _
                    ByVal UserID As String, _
                    ByVal Password As String) As Collection
End Function

Public Function GetTablesSql() As String
End Function

Public Function GetOptions() As clsImportOptions
End Function

Public Function CreateConnectionString(ByVal Provider As String, _
                    ByVal Server As String, _
                    ByVal User As String, _
                    ByVal Password As String, _
                    ByVal Database As String) As String
End Function

Public Function GetLogicalTable(conn As ADODB.Connection, TableName As String) As clsLogicalTable
End Function

