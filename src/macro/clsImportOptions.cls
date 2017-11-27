VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsImportOptions"
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

Public ConnectionMode           As enmConnectionMode
Public Provider                 As String
Public Database                 As String
Public DataSource               As String
Public DataSourceTip            As String
Public UserName                 As String
Public Password                 As String
Public connectionString         As String

Public ImportMode               As enmImportMode
Public ClearDataInExistedSheet  As Boolean

Public LastAccessTableName      As String
Public LastDatabaseName         As String

Private Sub Class_Initialize()
    ConnectionMode = ConnectionModeDataSource
    Provider = ""
    DataSource = ""
    
    ImportMode = ImportModeOverwrite
    ClearDataInExistedSheet = True
End Sub
