Attribute VB_Name = "basFile"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2012, Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit

Private Const ForReading As Integer = 1         '- Open a file for reading only. You can't write to this file.
Private Const ForWriting As Integer = 2
Private Const ForAppending As Integer = 8       '- Open a file and write to the end of the file.

Private Const TristateUseDefault As Integer = 2 '-- Opens the file using the system default.
Private Const TristateTrue As Integer = 1       '-- Opens the file as Unicode.
Private Const TristateFalse As Integer = 0      '-- Opens the file as ASCII.

Public Function CopyFolder(Source As String, Destination As String)
    Dim objFso As Object
    
    Set objFso = CreateObject("Scripting.FileSystemObject")
    objFso.CopyFolder objFso.GetAbsolutePathName(Source), objFso.GetAbsolutePathName(Destination)

End Function

Public Function FolderExists(ByRef Path)
    Dim objFso As Object
    
    Set objFso = CreateObject("Scripting.FileSystemObject")
    FolderExists = objFso.FolderExists(Path)

End Function

Public Function ReadFromTextFile(fileName As String) As String
    Dim Content As String
    Dim objStream
    
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.LoadFromFile fileName
    Content = objStream.ReadText()
    objStream.Close

    ReadFromTextFile = Content
End Function

Public Function ReadFromTextFileFso(fileName As String) As String
    Dim objFso As Object
    Dim objTextStream As Object
    Dim Content As String
    
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Set objTextStream = objFso.OpenTextFile(fileName, ForReading, TristateTrue)
    Content = objTextStream.ReadAll
    objTextStream.Close
    
    ReadFromTextFileFso = Content
End Function

Public Function SaveToTextFile(fileName As String, Content As String)
    Dim adTypeBinary: adTypeBinary = 1
    Dim adTypeText: adTypeText = 2
    Dim adSaveCreateOverWrite: adSaveCreateOverWrite = 2
    
    Dim objStream As Object
    Dim objStreamNoBOM As Object
    
    Set objStream = CreateObject("ADODB.Stream")
    Set objStreamNoBOM = CreateObject("ADODB.Stream")

    With objStream
        .Type = adTypeText
        .Open
        .Charset = "UTF-8"
        .WriteText Content
        .Position = 0
        .Type = adTypeBinary
        .Position = 3
   End With

   With objStreamNoBOM
      .Type = adTypeBinary
      .Open
      .Write objStream.Read
      .SaveToFile fileName, adSaveCreateOverWrite
      .Close
   End With

   objStream.Close

End Function

Public Sub DeleteFolder(Path As String)
    Dim objFso As Object
    Set objFso = CreateObject("Scripting.FileSystemObject")
    
    If objFso.FolderExists(Path) Then
        objFso.DeleteFolder Path, True
    End If
End Sub

Public Function MakeFolder(Path As String)
    If FolderExists(Path) Then
        Exit Function
    End If
    
    VBA.MkDir Path
End Function

Public Function SaveToTextFileFso(fileName As String, Content As String)
    Dim objFso As Object
    Dim objFile As Object
    
    Set objFso = CreateObject("Scripting.FileSystemObject")
'    If objFso.FileExists(FileName) Then
'        objFso.DeleteFile FileName, True
'    End If
    Set objFile = objFso.CreateTextFile(fileName, True, True)

    objFile.Write Content
    objFile.Close
End Function

Public Function Zip(ZipApp As String, _
                        SourceFolder As String, _
                        TargetFile As String)
                        
    Dim objFso As Object
    Set objFso = CreateObject("Scripting.FileSystemObject")
    
    RunExe ZipApp, "a -tzip " _
                    & """" & objFso.GetAbsolutePathName(TargetFile) & """" _
                    & " -r " _
                    & """" & objFso.GetAbsolutePathName(SourceFolder) & """"
                        
End Function

Private Sub RunExe(ExecuteFile As String, Arguments As String)
    Dim WshShell, oExec
    Set WshShell = CreateObject("WScript.Shell")

    Set oExec = WshShell.Exec("""" & ExecuteFile & """ " & Arguments)

    Do While oExec.Status = 0
        Application.Wait 500
    Loop
End Sub
