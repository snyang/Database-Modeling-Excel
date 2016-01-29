'------------------------------------------------------------------------------------------------
' Usage:
'------------------------------------------------------------------------------------------------

' Force explicit declaration of all variables.
Option Explicit

Dim verbose
verbose = true

Dim fsoStandard, console
Set fsoStandard = CreateObject ("Scripting.FileSystemObject")
Set console = fsoStandard.GetStandardStream (1)
'------------------------------------------------------------------------------------------------
'-- Declare Current Version
Dim sMajorVersion, sMinorVersion, sRevisionNumber, sReleaseType
'-- MajorVersion is changed for big feature be added.
sMajorVersion = "5"
'-- MinorVersion is changed for normal feature be added.
sMinorVersion = "0"
'-- Revision number is changed for bug fix.
sRevisionNumber = "1"

'-- Release type definition. a: arlfa release; b#: beta release; rc#:realease cadidate; <empty>: product release
sReleaseType = ""
Dim sVersion
sVersion = sMajorVersion _
            & "." & sMinorVersion _
            & "." & sRevisionNumber 
If Len(sReleaseType) <> 0 Then
    sVersion = sVersion  & " " & sReleaseType
End if
sVersion = Replace(sVersion, " ", "_")
sVersion = Replace(sVersion, ".", "_")

'-- Declare Constants 
Dim sSourcePath, sResourcePath
sSourcePath = "..\..\"
sResourcePath = "..\Resources\"

Dim sOutputFolder, sMacroFolder, sOutput_ResourceFolder, sDeployFolder
sOutputFolder = "..\00_output"
sMacroFolder = sOutputFolder & "\macros\" 
sOutput_ResourceFolder = sOutputFolder & "\DB_Modeling_Excel\"
sDeployFolder = sOutputFolder & "\deploy\"

Dim sDeployFile
sDeployFile = sDeployFolder & "database_modeling_template_" & sVersion & ".zip"

Dim sSourceExcelFilename

sSourceExcelFilename = sSourcePath & "_DatabaseModeling_Template_Source.xls"

'-- SQL Server
Dim sTemplate_SQLServer, sTemplate_SQLServer_Fullname, sOutput_Template_SQLServer_Fullname
sTemplate_SQLServer = "DatabaseModeling_Template_SQLServer"
sTemplate_SQLServer_Fullname = sOutput_ResourceFolder & sTemplate_SQLServer & ".xls"
sOutput_Template_SQLServer_Fullname = sOutput_ResourceFolder & sTemplate_SQLServer & "_" & sVersion & ".xls"

'-- MySQL
Dim sTemplate_MySQL, sTemplate_MySQL_Fullname, sOutput_Template_MySQL_Fullname
sTemplate_MySQL = "DatabaseModeling_Template_MySQL"
sTemplate_MySQL_Fullname = sOutput_ResourceFolder & sTemplate_MySQL & ".xls"
sOutput_Template_MySQL_Fullname = sOutput_ResourceFolder & sTemplate_MySQL & "_" & sVersion & ".xls"

'-- Oracle
Dim sTemplate_Oracle, sTemplate_Oracle_Fullname, sOutput_Template_Oracle_Fullname
sTemplate_Oracle = "DatabaseModeling_Template_Oracle"
sTemplate_Oracle_Fullname = sOutput_ResourceFolder & sTemplate_Oracle & ".xls"
sOutput_Template_Oracle_Fullname = sOutput_ResourceFolder & sTemplate_Oracle & "_" & sVersion & ".xls"

'-- SQLite
Dim sTemplate_SQLite, sTemplate_SQLite_Fullname, sOutput_Template_SQLite_Fullname
sTemplate_SQLite = "DatabaseModeling_Template_SQLite"
sTemplate_SQLite_Fullname = sOutput_ResourceFolder & sTemplate_SQLite & ".xls"
sOutput_Template_SQLite_Fullname = sOutput_ResourceFolder & sTemplate_SQLite & "_" & sVersion & ".xls"

'-- PostgreSQL
Dim sTemplate_PostgreSQL, sTemplate_PostgreSQL_Fullname, sOutput_Template_PostgreSQL_Fullname
sTemplate_PostgreSQL = "DatabaseModeling_Template_PostgreSQL"
sTemplate_PostgreSQL_Fullname = sOutput_ResourceFolder & sTemplate_PostgreSQL & ".xls"
sOutput_Template_PostgreSQL_Fullname = sOutput_ResourceFolder & sTemplate_PostgreSQL & "_" & sVersion & ".xls"

'-- Declare Constants for compress
Dim sZipApp
sZipApp = "C:\Program Files (x86)\7-Zip\7z.exe"

'------------------------------------------------------------------------------------------------
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

'  delete ouput folder
console.WriteLine  "Deleting ouput folder..."
If (fso.FolderExists(sOutputFolder)) Then
    Call fso.DeleteFolder(sOutputFolder, true)
End If

' create output folder
console.WriteLine  "Creating output folders..."
Dim f
Set f = fso.CreateFolder(sOutputFolder)
Set f = fso.CreateFolder(sMacroFolder)
Set f = fso.CreateFolder(sOutput_ResourceFolder)
Set f = fso.CreateFolder(sDeployFolder)

' export macros
console.WriteLine  "Exporting all macro files..."
RunVbs "runExcelMacro.vbs", "-f """ & fso.GetAbsolutePathName(sSourceExcelFilename) & """ -m ""VBComponent_ExportAll_Command"" -a """ & fso.GetAbsolutePathName(sMacroFolder) & """"

' copy resouces
console.WriteLine  "Copying templates..."
fso.CopyFolder fso.GetAbsolutePathName(sResourcePath), fso.GetAbsolutePathName(sOutput_ResourceFolder)
fso.MoveFile fso.GetAbsolutePathName(sTemplate_SQLServer_Fullname), fso.GetAbsolutePathName(sOutput_Template_SQLServer_Fullname)
fso.MoveFile fso.GetAbsolutePathName(sTemplate_MySQL_Fullname), fso.GetAbsolutePathName(sOutput_Template_MySQL_Fullname)
fso.MoveFile fso.GetAbsolutePathName(sTemplate_Oracle_Fullname), fso.GetAbsolutePathName(sOutput_Template_Oracle_Fullname)
fso.MoveFile fso.GetAbsolutePathName(sTemplate_SQLite_Fullname), fso.GetAbsolutePathName(sOutput_Template_SQLite_Fullname)
fso.MoveFile fso.GetAbsolutePathName(sTemplate_PostgreSQL_Fullname), fso.GetAbsolutePathName(sOutput_Template_PostgreSQL_Fullname)

'-- import macros into templates (and changes some variables)
console.WriteLine  "Updating SQL Server template..."
'-- Remvoe all method always failure.
'RunVbs "runExcelMacro.vbs", "-f """ & fso.GetAbsolutePathName(sOutput_Template_SQLServer_Fullname) & """ -m ""VBComponent_RemoveAll_Command"" -s"
RunVbs "runExcelMacro.vbs", "-f """ & fso.GetAbsolutePathName(sOutput_Template_SQLServer_Fullname) & """ -m ""VBComponent_ImportAll_Command"" -a """ & fso.GetAbsolutePathName(sMacroFolder) & """ -s"
RunVbs "runExcelMacro.vbs", "-f """ & fso.GetAbsolutePathName(sOutput_Template_SQLServer_Fullname) & """ -m ""ConfigureTheExcel"" -a ""SQL Server"" -s"

console.WriteLine  "Updating MySQL template..."
'RunVbs "runExcelMacro.vbs", "-f """ & fso.GetAbsolutePathName(sOutput_Template_MySQL_Fullname) & """ -m ""VBComponent_RemoveAll_Command"" -s"
RunVbs "runExcelMacro.vbs", "-f """ & fso.GetAbsolutePathName(sOutput_Template_MySQL_Fullname) & """ -m ""VBComponent_ImportAll_Command"" -a """ & fso.GetAbsolutePathName(sMacroFolder) & """ -s"
RunVbs "runExcelMacro.vbs", "-f """ & fso.GetAbsolutePathName(sOutput_Template_MySQL_Fullname) & """ -m ""ConfigureTheExcel"" -a ""MySQL"" -s"

console.WriteLine  "Updating Oracle template..."
RunVbs "runExcelMacro.vbs", "-f """ & fso.GetAbsolutePathName(sOutput_Template_Oracle_Fullname) & """ -m ""VBComponent_ImportAll_Command"" -a """ & fso.GetAbsolutePathName(sMacroFolder) & """ -s"
RunVbs "runExcelMacro.vbs", "-f """ & fso.GetAbsolutePathName(sOutput_Template_Oracle_Fullname) & """ -m ""ConfigureTheExcel"" -a ""Oracle"" -s"

console.WriteLine  "Updating SQLite template..."
RunVbs "runExcelMacro.vbs", "-f """ & fso.GetAbsolutePathName(sOutput_Template_SQLite_Fullname) & """ -m ""VBComponent_ImportAll_Command"" -a """ & fso.GetAbsolutePathName(sMacroFolder) & """ -s"
RunVbs "runExcelMacro.vbs", "-f """ & fso.GetAbsolutePathName(sOutput_Template_SQLite_Fullname) & """ -m ""ConfigureTheExcel"" -a ""SQLite"" -s"

console.WriteLine  "Updating PostgreSQL template..."
RunVbs "runExcelMacro.vbs", "-f """ & fso.GetAbsolutePathName(sOutput_Template_PostgreSQL_Fullname) & """ -m ""VBComponent_ImportAll_Command"" -a """ & fso.GetAbsolutePathName(sMacroFolder) & """ -s"
RunVbs "runExcelMacro.vbs", "-f """ & fso.GetAbsolutePathName(sOutput_Template_PostgreSQL_Fullname) & """ -m ""ConfigureTheExcel"" -a ""PostgreSQL"" -s"

'-- compress templates
console.WriteLine  "Zipping the packge..."
RunExe sZipApp, "a -tzip """ & fso.GetAbsolutePathName(sDeployFile) & """" _
                             & " -r " & fso.GetAbsolutePathName(sOutput_ResourceFolder)

Display "Done."

Sub RunExe(fileName, arguments)
    Dim WshShell, oExec
    Set WshShell = CreateObject("WScript.Shell")

    Set oExec = WshShell.Exec("""" & fileName & """ " & arguments)

    Do While oExec.Status = 0
        WScript.Sleep 100
    Loop
End Sub

Sub RunVbs(fileName, arguments)
    Dim WshShell
    Set WshShell = CreateObject("WScript.Shell")

    Call WshShell.Run(fileName & " " & arguments, 0, true)
End Sub

Sub Display(Msg)
	WScript.Echo Now & ". Error Code: " & Hex(Err) & " - " & Msg
End Sub

Sub Trace(Msg)
	if verbose = true then
		WScript.Echo Now & " : " & Msg	
	end if
End Sub