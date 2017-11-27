Attribute VB_Name = "basExcel"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2017, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

Private statusBarState As Boolean

Public Function DisableUI()
    
    'Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    'statusBarState = Application.DisplayStatusBar
    'Application.DisplayStatusBar = False
    'Application.EnableEvents = False

End Function

Public Function EnableUI()
    'Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    'Application.DisplayStatusBar = statusBarState
    'Application.EnableEvents = True

End Function

Public Sub RenderRangeMerge(ByRef Range As Range)
    Range.Merge
    With Range
        .HorizontalAlignment = xlLeft
        .MergeCells = True
    End With
End Sub


