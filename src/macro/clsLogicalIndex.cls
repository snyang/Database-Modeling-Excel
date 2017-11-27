VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogicalIndex"
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

Private mIKName As String
Private mIKColumns As String
Private mIsUnique As Boolean
Private mIsClustered As Boolean

''' --------------------------------------------------------
''' <summary>
'''     Property: IKName (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get IKName() As String
        IKName = mIKName
End Property
Public Property Let IKName(Value As String)
        mIKName = Value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: IKColumns (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get IKColumns() As String
        IKColumns = mIKColumns
End Property
Public Property Let IKColumns(Value As String)
        mIKColumns = Value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: IsUnique (Boolean)
''' </summary>
''' --------------------------------------------------------
Public Property Get IsUnique() As Boolean
        IsUnique = mIsUnique
End Property
Public Property Let IsUnique(Value As Boolean)
        mIsUnique = Value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: IsClustered (Boolean)
''' </summary>
''' --------------------------------------------------------
Public Property Get IsClustered() As Boolean
        IsClustered = mIsClustered
End Property
Public Property Let IsClustered(Value As Boolean)
        mIsClustered = Value
End Property

