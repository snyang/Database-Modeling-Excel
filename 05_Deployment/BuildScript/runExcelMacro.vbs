'------------------------------------------------------------------------------------------------
'
' Usage: runExcelMacro <--filename|-f ExcelFilename>  <--macroname|-m macroName > 
'                                   [--arguments|-a arg1[,server2,server3...]]
'                                   [--verbose|-v]
'                                   [--help|-?]
'
'
' Example 1: runExcelMacro -f "E:\_MyFiles\Projets\DBModelExcel\DatabaseModeling_Template_2_0_a.xls" -m "VBComponent_ExportAll" -a "E:\_MyFiles\Projets\DBModelExcel\Build\macros"
'------------------------------------------------------------------------------------------------

' Force explicit declaration of all variables.
Option Explicit

On Error Resume Next

Dim oArgs, ArgNum
Dim sSourceFile
Dim sMacroName, aMacroArgs
Dim bNeedSave
bNeedSave = false
Dim verbose

verbose = false
Set oArgs = WScript.Arguments
ArgNum = 0

While ArgNum < oArgs.Count

	Select Case LCase(oArgs(ArgNum))
		Case "--filename", "-f":
			ArgNum = ArgNum + 1
			sSourceFile=oArgs(ArgNum)
		Case "--macroname", "-m":
			ArgNum = ArgNum + 1
			sMacroName=oArgs(ArgNum)
		Case "--arguments","-a":
			ArgNum = ArgNum + 1
			aMacroArgs = Split(oArgs(ArgNum), ",", -1)
		Case "--needSave", "-s":
			bNeedSave = true
		Case "--verbose", "-v":
			verbose = true
        Case "--help","-?":
			Call DisplayUsage
		Case Else:
			Call DisplayUsage
	End Select	

	ArgNum = ArgNum + 1
Wend

Dim oExcelApp
Set oExcelApp = CreateObject("Excel.Application")
oExcelApp.Visible = false
Dim oExcelBook
Set oExcelBook = oExcelApp.Workbooks.Open(sSourceFile)

Dim iMacroArgsCount
iMacroArgsCount = UBound(aMacroArgs)
'-- Display iMacroArgsCount
SELECT CASE iMacroArgsCount
CASE "":
    Call oExcelApp.Run(sMacroName)
CASE 0:
    Call oExcelApp.Run(sMacroName, aMacroArgs(0))
CASE 1:
    Call oExcelApp.Run(sMacroName, aMacroArgs(0), aMacroArgs(1))
CASE 2:
    Call oExcelApp.Run(sMacroName, aMacroArgs(0), aMacroArgs(1), aMacroArgs(2))
CASE 3:
    Call oExcelApp.Run(sMacroName, aMacroArgs(0), aMacroArgs(1), aMacroArgs(2), aMacroArgs(3))
CASE 4:
    Call oExcelApp.Run(sMacroName, aMacroArgs(0), aMacroArgs(1), aMacroArgs(2), aMacroArgs(3), aMacroArgs(4))
CASE ELSE
    Trace "not implement!"
END SELECT

If (bNeedSave) Then
    oExcelBook.Save()
End If
oExcelApp.Quit
 
Sub Display(Msg)
	'WScript.Echo Now & ". Error Code: " & Hex(Err) & " - " & Msg
    WScript.Echo Now & ". Message:" & Msg
End Sub

Sub Trace(Msg)
	if verbose = true then
		WScript.Echo Now & " : " & Msg	
	end if
End Sub
	
Sub DisplayUsage
	WScript.Echo "Usage: runExcelMacro <--filename|-f ExcelFilename>  <--macroname|-m macroName > "
	WScript.Echo "                     [--arguments|-a arg1[,server2,server3...]]"
	WScript.Echo "'                    [--verbose|-v]"
	WScript.Echo "                     [--help|-?]"
	WScript.Echo "Example 1: runExcelMacro -f ""E:\_MyFiles\Projets\DBModelExcel\DatabaseModeling_Template_2_0_a.xls"" -m ""VBComponent_ExportAll"" -a ""E:\_MyFiles\Projets\DBModelExcel\Build\macros"""
	WScript.Quit (1)
End Sub
