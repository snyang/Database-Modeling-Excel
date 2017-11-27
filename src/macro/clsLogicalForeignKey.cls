VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogicalForeignKey"
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

Private mFKName             As String
Private mFKColumns          As String
Private mRefTableName       As String
Private mRefTableColumns    As String
Private mFKOption           As String
Private mOnUpdate           As String
Private mOnDelete           As String

''' --------------------------------------------------------
''' <summary>
'''     Property: FKName (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get FKName() As String
        FKName = mFKName
End Property
Public Property Let FKName(Value As String)
        mFKName = Value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: FKColumns (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get FKcolumns() As String
        FKcolumns = mFKColumns
End Property
Public Property Let FKcolumns(Value As String)
        mFKColumns = Value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: RefTableAndColumns (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get RefTableAndColumns() As String
        RefTableAndColumns = mRefTableName & "(" & mRefTableColumns & ")"
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: RefTableName (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get RefTableName() As String
        RefTableName = mRefTableName
End Property
Public Property Let RefTableName(Value As String)
        mRefTableName = Value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: RefTableColumns (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get RefTableColumns() As String
        RefTableColumns = mRefTableColumns
End Property
Public Property Let RefTableColumns(Value As String)
        mRefTableColumns = Value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: FKOption (String)
'''     like "ON DELETE CASCADE"
''' </summary>
''' --------------------------------------------------------
Public Property Get fkOption() As String
        fkOption = mFKOption
End Property
Public Property Let fkOption(Value As String)
        mFKOption = TrimEx(Value)
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: OnUpdate (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get OnUpdate() As String
        OnUpdate = mOnUpdate
End Property
Public Property Let OnUpdate(Value As String)
        mOnUpdate = Value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: OnDelete (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get OnDelete() As String
        OnDelete = mOnDelete
End Property
Public Property Let OnDelete(Value As String)
        mOnDelete = Value
End Property


