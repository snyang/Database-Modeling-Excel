Attribute VB_Name = "basString"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2012, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

''' --------------------------------------------------------
''' <summary>
'''     Trim a string include space, vblf, vbcr.
''' </summary>
''' <param name="text"></param>
''' <returns>trimed string</returns>
''' <remarks>
''' </remarks>
''' --------------------------------------------------------
Public Function TrimEx(ByVal Text As String) As String
    Dim sRet As String
    sRet = Trim(Text)
    Do While Len(sRet) > 0
        If Left(sRet, 1) = vbLf Or Left(sRet, 1) = vbCr Then
            sRet = Mid(sRet, 2)
        Else
            Exit Do
        End If
    Loop
    Do While Len(sRet) > 0
        If Right(sRet, 1) = vbLf Or Right(sRet, 1) = vbCr Then
            sRet = Mid(sRet, 1, Len(sRet) - 1)
        Else
            Exit Do
        End If
    Loop

    '-- Return
    TrimEx = sRet
End Function

''' --------------------------------------------------------
''' <summary>
'''     Format string.
''' </summary>
''' <param name="text"></param>
''' <param name="Args"></param>
''' <returns></returns>
''' <remarks>
'''     format like:
'''     "a{0:description}b{1}c\{0\}"
'''     {0:description} is a tag; {1} is a tag; \{0\} is {0}, a,b,c is a,b,c
''' </remarks>
''' --------------------------------------------------------
Public Function FormatString(ByVal Text As String, ParamArray Args()) As String
    Dim newText             As String
    Dim index               As Long

    Dim textLength          As Long
    Dim argLength           As Integer
    Dim ch                  As String

    Dim tagBeginInNewText   As Long
    Dim tagText             As String
    Dim tagValue            As Integer
    Dim tagLength           As Long

    textLength = Len(Text)
    argLength = UBound(Args)
    index = 1

    tagBeginInNewText = -1
    tagText = ""
    tagValue = -1

    Do While index <= textLength
        ch = Mid(Text, index, 1)
        Select Case ch
            Case "\"
                '-- Escape char, escape "\", "{", "}"
                If index < textLength Then
                    If Mid(Text, index + 1, 1) = "\" _
                        Or Mid(Text, index + 1, 1) = "{" _
                        Or Mid(Text, index + 1, 1) = "}" Then
                        ch = Mid(Text, index + 1, 1)
                        index = index + 1
                        GoTo FLAG_AddToText
                    End If
                End If
            Case "{"
                tagBeginInNewText = Len(newText)
                tagText = ""
                tagValue = -1
            Case "}"
                If tagBeginInNewText >= 0 Then
                    If tagValue = -1 Then
                        tagLength = Len(tagText)
                        If tagLength > 0 And tagLength <= 4 Then
                            tagValue = CInt(tagText)
                            If tagValue > argLength Then
                                tagBeginInNewText = -1
                                tagText = ""
                                tagValue = -1
                            End If
                        Else
                            tagBeginInNewText = -1
                            tagText = ""
                            tagValue = -1
                        End If
                    End If
                    If tagValue >= 0 Then
                        newText = Mid(newText, 1, tagBeginInNewText) & Args(tagValue)
                        GoTo Flag_next
                    End If
                End If
            Case Else
                If tagBeginInNewText >= 0 Then
                    If IsNumeric(ch) And tagValue = -1 Then
                        tagText = tagText & ch
                    ElseIf ch = ":" Then
                        tagLength = Len(tagText)
                        If tagLength > 0 And tagLength <= 4 Then
                            tagValue = CInt(tagText)
                            If tagValue > argLength Then
                                tagBeginInNewText = -1
                                tagText = ""
                                tagValue = -1
                            End If
                        Else
                            tagBeginInNewText = -1
                            tagText = ""
                            tagValue = -1
                        End If
                    End If
                End If
        End Select
FLAG_AddToText:
        newText = newText & ch
Flag_next:
        index = index + 1
    Loop

    '-- Return
    FormatString = newText
End Function

''' --------------------------------------------------------
''' <summary>
'''     Get string before indicate string
''' </summary>
''' <param name="text"></param>
''' <param name="subString"></param>
''' <returns></returns>
''' <remarks>
'''     text = "table(col1, col2)"
'''     text2 = "("
'''     return "table"
''' </remarks>
''' --------------------------------------------------------
Public Function GetStringBefore(ByVal Text As String, ByVal Text2 As String) As String
    Dim pos     As Long
    pos = VBA.Strings.InStr(Text, Text2)
    If pos > 0 Then
        Text = Mid(Text, 1, pos - 1)
    End If
    
    GetStringBefore = Text
End Function

''' --------------------------------------------------------
''' <summary>
'''     Get string after indicate string
''' </summary>
''' <param name="text"></param>
''' <param name="subString"></param>
''' <returns></returns>
''' <remarks>
'''     text = "table(col1, col2"
'''     text2 = "("
'''     return "col1, col2"
''' </remarks>
''' --------------------------------------------------------
Public Function GetStringAfter(ByVal Text As String, ByVal Text2 As String) As String
    Dim pos     As Long
    pos = VBA.Strings.InStr(Text, Text2)
    If pos > 0 Then
        Text = Mid(Text, pos + 1)
    End If
    
    GetStringAfter = Text
End Function

''' --------------------------------------------------------
''' <summary>
'''     Reverse the specific string.
''' </summary>
''' <param name="Text"></param>
''' <returns></returns>
''' --------------------------------------------------------
Public Function Reverse(ByVal Text As String) As String
    Dim newText As String
    Dim index   As Integer
    Dim length  As Long
    
    length = Len(Text)
    For index = 1 To length
        newText = newText & Mid(Text, length - index + 1, 1)
    Next
    Reverse = newText
End Function
''' --------------------------------------------------------
''' <summary>
'''     split and trim
''' </summary>
''' <param name="text"></param>
''' <param name="delimiter"></param>
''' <returns></returns>
''' <remarks>
''' </remarks>
''' --------------------------------------------------------
Public Function SplitAndTrim(ByVal Text As String, ByVal delimiter As String) As String()
    Dim str()   As String
    Dim index   As Integer
    str = Split(Text, delimiter)
    
    For index = LBound(str) To UBound(str)
        str(index) = Trim(str(index))
    Next
    SplitAndTrim = str
End Function

Public Function StartWith(Text1 As String, Text2 As String) As Boolean
    StartWith = VBA.StrComp(Mid(Text1, 1, Len(Text2)), Text2, vbTextCompare) = 0
End Function

''' --------------------------------------------------------
''' <summary>
'''     strip brackets, used to refine default value.
''' </summary>
''' <param name="value"></param>
''' <returns>stripped values</returns>
''' <remarks>
''' </remarks>
''' --------------------------------------------------------
Public Function StripBrackets(Value As String, Optional stripAll As Boolean = False) As String
    Dim minLength As Integer
    
    minLength = IIf(stripAll, 2, 4)
    
    Do While Len(Value) > minLength
        If stripAll Then
            If Mid(Value, 1, 1) = "(" And Mid(Value, Len(Value), 1) = ")" Then
                Value = Mid(Value, 2, Len(Value) - 2)
            Else
                Exit Do
            End If
        Else
            If Mid(Value, 1, 2) = "((" And Mid(Value, Len(Value) - 1, 2) = "))" Then
                Value = Mid(Value, 2, Len(Value) - 2)
            Else
                Exit Do
            End If
        End If
    Loop
    
    StripBrackets = Value
End Function
