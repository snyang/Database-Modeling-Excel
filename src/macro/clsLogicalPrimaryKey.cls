VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogicalPrimaryKey"
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

Private mPKName As String
Private mPKColumns As String
Private mIsClustered As Boolean

''' --------------------------------------------------------
''' <summary>
'''     Property: PKName (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get PKName() As String
    PKName = mPKName
End Property
Public Property Let PKName(Value As String)
    mPKName = Value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: PKColumns (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get PKcolumns() As String
    PKcolumns = mPKColumns
End Property
Public Property Let PKcolumns(Value As String)
    mPKColumns = Value
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

Private Sub Class_Initialize()
    mIsClustered = True
End Sub
