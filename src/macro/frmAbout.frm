VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "About Database Modeling Excel"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2013, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

Private Sub btnDenote_Click()
 On Error GoTo Flag_Err
    Call VBA.Shell("cmd /C ""start http://sourceforge.net/donate/index.php?group_id=171489""")
    
    Exit Sub
Flag_Err:
    Call MsgBoxEx_Error
End Sub

Private Sub txtEmail_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo Flag_Err
    Call VBA.Shell("cmd /C ""start mailto:" & txtEmail & """")
    
    Exit Sub
Flag_Err:
    Call MsgBoxEx_Error
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "About " & App_Name
    Me.labName = App_Name & " < " & App_Version & " >"
    Me.txtLicense = "Copyright (c) 2013, Yang Ning (Steven)" _
        & vbCrLf & "All rights reserved." _
        & vbCrLf & "" _
        & vbCrLf & "Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:" _
        & vbCrLf & "" _
        & vbCrLf & "* Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer." _
        & vbCrLf & "" _
        & vbCrLf & "* Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.""" _
        & vbCrLf & "" _
        & vbCrLf & "THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS ""AS IS"" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE."
    Me.txtLicense.SelStart = 0
    
    Dim desc As String
    desc = "Support open source!" _
        & vbCrLf & "To contribute more software, make the world better." _
        & vbCrLf & "" _
        & vbCrLf & "What's a reasonable donation?" _
        & vbCrLf & "Even $1 is enough to show your appreciation - give what you can." _
        & vbCrLf & "" _
        & vbCrLf & "If you are a business, using " & App_Name & " for professional endeavors, you should consider donating at least $20. If you are using several copies of the program in your company you should perhaps pay an additional $5 for each copy. In the end, pay what you feel is fair."
        

End Sub
