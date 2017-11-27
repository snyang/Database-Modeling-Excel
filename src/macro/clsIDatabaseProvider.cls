VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsIDatabaseProvider"
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
'-----------------------------------------------------------
'-- Interface for databases provider
'-- which provides functions needed in the application.
'-----------------------------------------------------------

Public Sub GetSQLCreateTable(objLogicalTableCollection As Collection, _
                            ByVal withComment As Boolean, _
                            Optional outputOptions As clsOutputOptions)

End Sub

Public Sub GetSQLDropTable(objLogicalTableCollection As Collection, _
                            Optional outputOptions As clsOutputOptions)

End Sub

Public Sub GetSQLDropAndCreateTable(objLogicalTableCollection As Collection, _
                            ByVal withComment As Boolean, _
                            Optional outputOptions As clsOutputOptions)
End Sub

Public Sub GetSQLCreateTableIfNotExists(objLogicalTableCollection As Collection, _
                            Optional outputOptions As clsOutputOptions)

End Sub
