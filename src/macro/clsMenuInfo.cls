VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMenuInfo"
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

Private mParent         As String
Private mCaption        As String
Private mTooltipText    As String
Private mOnAction       As String
Private mStyle          As MsoControlType


Private mInstance       As Object
Private mInstanceIndex  As Integer
Private mChildCount     As Integer

''' --------------------------------------------------------
''' <summary>
'''     Property: Parent (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get Parent() As String
        Parent = mParent
End Property
Public Property Let Parent(Value As String)
        mParent = Value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: Caption (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get Caption() As String
        Caption = mCaption
End Property
Public Property Let Caption(Value As String)
        mCaption = Value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: TooltipText (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get TooltipText() As String
        TooltipText = mTooltipText
End Property
Public Property Let TooltipText(Value As String)
        mTooltipText = Value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: OnAction (String)
''' </summary>
''' --------------------------------------------------------
Public Property Get OnAction() As String
        OnAction = mOnAction
End Property
Public Property Let OnAction(Value As String)
        mOnAction = Value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: Style (MsoControlType)
''' </summary>
''' --------------------------------------------------------
Public Property Get Style() As MsoControlType
        Style = mStyle
End Property
Public Property Let Style(Value As MsoControlType)
        mStyle = Value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: Instance (Object)
''' </summary>
''' --------------------------------------------------------
Public Property Get Instance() As Object
        Set Instance = mInstance
End Property
Public Property Set Instance(Value As Object)
        Set mInstance = Value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: InstanceIndex (Integer)
''' </summary>
''' --------------------------------------------------------
Public Property Get InstanceIndex() As Integer
        InstanceIndex = mInstanceIndex
End Property
Public Property Let InstanceIndex(Value As Integer)
        mInstanceIndex = Value
End Property

''' --------------------------------------------------------
''' <summary>
'''     Property: ChildCount (Integer)
''' </summary>
''' --------------------------------------------------------
Public Property Get ChildCount() As Integer
        ChildCount = mChildCount
End Property
Public Property Let ChildCount(Value As Integer)
        mChildCount = Value
End Property

