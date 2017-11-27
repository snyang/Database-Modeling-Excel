Attribute VB_Name = "basShell"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2014, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

Private Const WindowStyle As Integer = 1                '- Window Style
Private Const WaitOnReturn As Boolean = True            '- Wait the command return

Public Function RunCommand(Command As String) As Integer
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim returnCode As Integer
    
    returnCode = wsh.Run(Command, WindowStyle, WaitOnReturn)
    RunCommand = returnCode
End Function
