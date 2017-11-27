VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOutputOptions"
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

Private mOutputMode     As enmOutputMode
Private mOutputFullName As String

Public Property Get OutputMode() As enmOutputMode
    OutputMode = mOutputMode
End Property
Public Property Let OutputMode(Value As enmOutputMode)
    mOutputMode = Value
End Property

Public Property Get OutputFullName() As String
    OutputFullName = mOutputFullName
End Property
Public Property Let OutputFullName(Value As String)
    mOutputFullName = Value
End Property

