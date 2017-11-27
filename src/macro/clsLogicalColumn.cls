VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogicalColumn"
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

Private mColumnLabel    As String
Private mColumnName     As String
Private mDataType       As String
Private mNullable       As Boolean
Private mDefault        As String
Private mComment        As String

Public Property Get ColumnLabel() As String
    ColumnLabel = mColumnLabel
End Property
Public Property Let ColumnLabel(Value As String)
    mColumnLabel = Value
End Property

Public Property Get ColumnName() As String
    ColumnName = mColumnName
End Property
Public Property Let ColumnName(Value As String)
    mColumnName = Value
End Property

Public Property Get DataType() As String
    DataType = mDataType
End Property
Public Property Let DataType(Value As String)
    mDataType = Value
End Property

Public Property Get Nullable() As Boolean
    Nullable = mNullable
End Property
Public Property Let Nullable(Value As Boolean)
    mNullable = Value
End Property
    
Public Property Get Default() As String
    Default = mDefault
End Property
Public Property Let Default(Value As String)
    mDefault = Value
End Property

Public Property Get Comment() As String
    Comment = mComment
End Property
Public Property Let Comment(Value As String)
    mComment = Value
End Property

Public Property Get DBDescription() As String
    Const DBLabelFormat = "(Label: {0: label})"
    DBDescription = Me.Comment
    
    If Len(Me.ColumnLabel) > 0 Then
        If Len(DBDescription) > 0 Then
            DBDescription = DBDescription & " "
        End If
        DBDescription = DBDescription & FormatString(DBLabelFormat, Me.ColumnLabel)
    End If
    
    DBDescription = Trim(DBDescription)
End Property
Public Property Let DBDescription(Value As String)
    Me.Comment = Value
    
    Value = Reverse(Value)
    Dim newLabel As String
    newLabel = basString.GetStringBefore(Value, "(")
    If Len(newLabel) Then
        newLabel = Reverse(newLabel)
        If Mid(newLabel, 1, 6) = "Label:" And Mid(newLabel, Len(newLabel), 1) = ")" Then
            Me.Comment = Reverse(basString.GetStringAfter(Value, "("))
            Me.ColumnLabel = Trim(basString.GetStringBefore(Mid(newLabel, 7), ")"))
        End If
    End If
End Property
