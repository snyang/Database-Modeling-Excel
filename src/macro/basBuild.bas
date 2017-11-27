Attribute VB_Name = "basBuild"
'===========================================================
'-- Database Modeling Excel
'===========================================================
'-- Copyright (c) 2014, 2017 Yang Ning (Steven)
'-- All rights reserved.
'-- Email: steven.n.yang@gmail.com
'===========================================================
Option Explicit
Private Const ZipApp                As String = "C:\Program Files\7-Zip\7z.exe"

Private Const PackageNameFormat     As String = "Database_Modeling_Excel_{0: app version}"
Private Const TestDataLocation      As String = "\..\test\TestData"
Private Const RulesFileLocation     As String = "\data\DB_Rules.xlsx"
Private Const TableDataFileLocation As String = "\data\DB_TableDefinitions.xlsx"
Private Const TestExpectedFileName  As String = "\Test_Expected_{0}.sql"
Private Const TestGeneratedFileName As String = "\Test_Generated_{0}.sql"

Private PackageName             As String
Private MacroLocation           As String
Private ResourceLocation        As String
Private OutputLocation          As String
Private BuildLocation           As String
Private BuildDeployLocation     As String
Private BuildTestLocation       As String
Private IsDebug                 As Boolean


Public Function Build()
    On Error GoTo Flag_Err
    
    Dim StartTime As Date
    StartTime = Now
    
    '-- Debug flag
    IsDebug = False
    
    If Not IsDebug Then
        DisableUI
    End If
    
    PackageName = FormatString(PackageNameFormat, Replace(App_Version, ".", "_"))
    MacroLocation = ThisWorkbook.Path & "\macro"
    ResourceLocation = ThisWorkbook.Path & "\resources"
    OutputLocation = ThisWorkbook.Path & "\..\release"
    BuildLocation = ThisWorkbook.Path & "\..\temp"
    BuildDeployLocation = BuildLocation & "\" & PackageName
    BuildTestLocation = BuildLocation & "\test"
    
    basFile.DeleteFolder OutputLocation
    basFile.MakeFolder OutputLocation

    basFile.DeleteFolder BuildLocation
    basFile.MakeFolder BuildLocation
    basFile.MakeFolder BuildDeployLocation
    basFile.MakeFolder BuildTestLocation
    
    GenerateMacroFiles
    
    GenerateWorkbooks

    CopyResources
    
    Packge
    
    EnableUI
    Dim EndTime As Date
    EndTime = Now
    
    Dim elapseTime As Long
    elapseTime = VBA.DateDiff("s", StartTime, EndTime)
    MsgBox "Build Done in (" & elapseTime & "s)!" & IIf(IsDebug, " Debug Mode!", ""), vbInformation, App_Name
    Exit Function
    
Flag_Err:
    EnableUI
    Application.DisplayAlerts = True
    Call MsgBoxEx_Error
End Function

Private Sub CopyResources()
    basFile.CopyFolder ResourceLocation, BuildDeployLocation
End Sub

Private Sub Packge()
    Dim deployFileName     As String
    deployFileName = OutputLocation & "\" & PackageName & ".zip"
    
    basFile.Zip ZipApp, BuildDeployLocation, deployFileName
End Sub

Private Sub GenerateMacroFiles()
    
    basFile.DeleteFolder MacroLocation
    basFile.MakeFolder MacroLocation
    
    '-- copy macro
    mdlExcelFunctions.VBComponent_ExportAll_Command MacroLocation
    
End Sub


Private Sub GenerateWorkbooks()
    Dim dbType As String
    Dim objWorkbook As Workbook
    Dim DeployedPath As String
    DeployedPath = BuildDeployLocation
    
    Dim dbAllTypes(8) As String
    If IsDebug Then
        dbAllTypes(1) = basAppSetting.DBName_DB2
    Else
        dbAllTypes(1) = basAppSetting.DBName_DB2
        dbAllTypes(2) = basAppSetting.DBName_MariaDB
        dbAllTypes(3) = basAppSetting.DBName_MySQL
        dbAllTypes(4) = basAppSetting.DBName_Oracle
        dbAllTypes(5) = basAppSetting.DBName_PostgreSQL
        dbAllTypes(6) = basAppSetting.DBName_SQLite
        dbAllTypes(7) = basAppSetting.DBName_SQLServer
        dbAllTypes(8) = basAppSetting.DBName_All
    End If
    
    Dim index As Integer
    For index = 1 To UBound(dbAllTypes)
        dbType = dbAllTypes(index)
        
        If Len(dbType) = 0 Then
            Exit For
        End If
        
        '-- create workbook
        Set objWorkbook = Application.Workbooks.Add
        
        '-- copy macro
        VBComponent_CopyAll_Command ThisWorkbook, objWorkbook
        
        '-- create sheets
        EmptyWorkbook objWorkbook
        BuildSheetHistory objWorkbook
        BuildSheetRules objWorkbook, IIf(dbType = DBName_All, DBName_SQLServer, dbType)
        BuildSheetTables objWorkbook, IIf(dbType = DBName_All, DBName_SQLServer, dbType)
        BuildSheetIndex objWorkbook
        Application.Run objWorkbook.Name & "!Command_SetSheetsName_Click"
        Application.Run objWorkbook.Name & "!ConfigureTheExcel", dbType
        
        '-- configure print pages
        Call SetPageSetup(objWorkbook)
        
        objWorkbook.Activate
        objWorkbook.Sheets(1).Select
        
        '-- Save
        Dim workbookName As String
        workbookName = basString.FormatString("\DME_Template_{0: DB type}_{1: app version}.xlsm", _
                        Replace(dbType, " ", ""), _
                        Replace(App_Version, ".", "_"))
        
        Dim workbookPath As String
        Dim workbookBuildPath As String
        workbookPath = DeployedPath & workbookName
        workbookBuildPath = DeployedPath & Replace(workbookName, ".xlsm", "_building.xlsm")
        
        Application.DisplayAlerts = False
        objWorkbook.SaveAs workbookBuildPath, _
                        FileFormat:=xlOpenXMLWorkbookMacroEnabled, _
                        CreateBackup:=False
        objWorkbook.Close
        Application.DisplayAlerts = True
        
        '-- Test
        If dbType <> basAppSetting.DBName_All Then
            RunUnitTest workbookBuildPath, dbType
        End If
    
        Name workbookBuildPath As workbookPath
    Next
End Sub

Private Sub RunUnitTest(workbookPath As String, dbType As String)

    Dim targetPath As String
    Dim expectPath As String
    targetPath = BuildTestLocation & basString.FormatString(TestGeneratedFileName, Replace(dbType, " ", ""))
    expectPath = ThisWorkbook.Path & TestDataLocation & basString.FormatString(TestExpectedFileName, Replace(dbType, " ", ""))
    
    Dim objWorkbook As Workbook
    Set objWorkbook = Application.Workbooks.Open(workbookPath)
    
    '-- Generate SQL for testing
    Application.Run objWorkbook.Name & "!SaveAllDropAndCreateTableWithDescriptionSQL", targetPath
    objWorkbook.Close False
    
    Dim targetContent As String
    Dim expectContent As String
    targetContent = ReadFromTextFile(targetPath)
    expectContent = ReadFromTextFile(expectPath)
    
    If targetContent <> expectContent Then
        MsgBox basString.FormatString("Test Failed: database type '{0}'", dbType), vbCritical, App_Name
    End If
    
End Sub

Private Sub EmptyWorkbook(objWorkbook As Workbook)
    Dim objSheet As Worksheet
    
    While objWorkbook.Sheets.Count > 1
        Set objSheet = objWorkbook.Sheets(2)
        objSheet.Delete
    Wend
End Sub

Private Function BuildSheetIndex(objWorkbook As Workbook) As Worksheet
    Dim objSheet As Worksheet
    Set objSheet = objWorkbook.Sheets(1)
    objSheet.Activate
    objSheet.Name = basAppSetting.Sheet_Index_Name
    Call SetStyleSheet(objSheet)
    
    '-- Title
    objSheet.Cells(1, "C").Value = App_Name
    objSheet.Rows("1:1").RowHeight = 30
    objSheet.Cells(1, "C").Font.Size = 20
    objSheet.Cells(1, "C").Font.Bold = True
    objSheet.Cells(1, "C").HorizontalAlignment = xlCenter
    
    '-- Header
    Dim headerRow As Integer
    headerRow = 3
    objSheet.Cells(headerRow, "A").Value = "No"
    objSheet.Cells(headerRow, "C").Value = "Table"
    objSheet.Cells(headerRow, "D").Value = "Page"
    objSheet.Cells(headerRow, "E").Value = "..."
    Call RenderHeader(objSheet, "A" & headerRow & ":E" & headerRow)
    
    '-- Rows
    Dim index As Integer
    Dim iRow As Integer
    Dim objCell As Range
    
    For index = 1 To objWorkbook.Sheets.Count - 1
        iRow = index + headerRow
        If index = 1 Then
            '-- First Row
            objSheet.Cells(iRow, "A").Value = "1"
            objSheet.Cells(iRow, "D").Value = "4"
        Else
        
            objSheet.Cells(iRow, "A").Value = "=A" & iRow - 1 & "+1"
            objSheet.Cells(iRow, "D").Value = "=IF(C" & iRow & " = """", """", D" & iRow - 1 & " + IF( INT(E" & iRow - 1 & ") = 0, 1, E" & iRow - 1 & ") )"
        End If
        
        objSheet.Cells(iRow, "C").Value = "= GetTableName(A" & iRow & "+1)"
        Set objCell = objSheet.Cells(iRow, "B")
        objCell.Value = ">>"
        objCell.Hyperlinks.Delete
        If index <= objWorkbook.Sheets.Count Then
            objSheet.Hyperlinks.Add objCell, _
                    "", _
                    "'" & GetTableNameFromSheet(objWorkbook.Sheets(index + 1)) & "'!A1"
            RenderRangeLink objCell
        End If
    Next
    
    '-- Format
    objSheet.Columns("A:A").ColumnWidth = 4
    objSheet.Columns("B:B").ColumnWidth = 4
    objSheet.Columns("C:C").ColumnWidth = 100
    objSheet.Columns("D:D").ColumnWidth = 4
    objSheet.Columns("E:E").ColumnWidth = 4
    
    '-- Return
    objSheet.Activate
    objSheet.Cells(1, 1).Select
    Set BuildSheetIndex = objSheet
End Function

Private Function BuildSheetHistory(objWorkbook As Workbook) As Worksheet
    Dim objSheet As Worksheet
    Set objSheet = objWorkbook.Sheets.Add(, objWorkbook.Sheets(objWorkbook.Sheets.Count))
    objSheet.Name = basAppSetting.Sheet_Update_History_Name
    Call SetStyleSheet(objSheet)
    
    '-- Header
    Dim headerRow As Integer
    headerRow = 1
    objSheet.Cells(headerRow, "A").Value = "Date"
    objSheet.Cells(headerRow, "B").Value = "Author"
    objSheet.Cells(headerRow, "C").Value = "Sheet"
    objSheet.Cells(headerRow, "D").Value = "Comments"
    
    '-- Format
    RenderTableHeader objSheet, "A" & headerRow & ":E" & headerRow
    RenderTableBorder objSheet.Range("A" & headerRow & ":E" & headerRow + 10)
    objSheet.Columns("A:A").ColumnWidth = 4
    objSheet.Columns("B:B").ColumnWidth = 12
    objSheet.Columns("C:C").ColumnWidth = 12
    objSheet.Columns("D:D").ColumnWidth = 24
    objSheet.Columns("E:E").ColumnWidth = 64
    
    '-- Return
    objSheet.Cells(1, 1).Select
    Set BuildSheetHistory = objSheet
End Function

Private Function BuildSheetRules(objWorkbook As Workbook, _
                                    dbType As String) As Worksheet
    Dim objSheet As Worksheet
    Set objSheet = objWorkbook.Sheets.Add(, objWorkbook.Sheets(objWorkbook.Sheets.Count))
    objSheet.Name = basAppSetting.Sheet_Rule_Name
    Call SetStyleSheet(objSheet)
    
    objSheet.Cells(1, "A").Value = "Rules"
    Call RenderRangeBold(objSheet.Cells(1, "A"))
    
    Dim bookRules As Workbook
    Set bookRules = Application.Workbooks.Open(ThisWorkbook.Path & RulesFileLocation)
    Dim sheetRule As Worksheet
    Set sheetRule = bookRules.Worksheets(dbType)
    
    Dim iRow As Integer
    Dim currentRule As String
    For iRow = 2 To sheetRule.UsedRange.Rows.Count
        If Len(sheetRule.Cells(iRow, "A").Value) > 0 Then
            objSheet.Cells(iRow, "A").Value = "-"
            Call RenderRangeBold(objSheet.Cells(iRow, "A"))
            currentRule = sheetRule.Cells(iRow, "A").Value
        End If
        Call RenderRangeCenter(objSheet.Cells(iRow, "A"))
        
        If Len(sheetRule.Cells(iRow, "A").Value) Then
            objSheet.Cells(iRow, "B").Value = sheetRule.Cells(iRow, "A").Value
            Call RenderRangeBold(objSheet.Cells(iRow, "B"))
        End If
        If Len(sheetRule.Cells(iRow, "B").Value) Then
            objSheet.Cells(iRow, "C").Value = sheetRule.Cells(iRow, "B").Value
        End If
        If Len(sheetRule.Cells(iRow, "C").Value) Then
            objSheet.Cells(iRow, "D").Value = sheetRule.Cells(iRow, "C").Value
            Call RenderRuleSample(currentRule, objSheet.Cells(iRow, "D"))
        End If
        If Len(sheetRule.Cells(iRow, "D").Value) Then
            objSheet.Cells(iRow, "E").Value = sheetRule.Cells(iRow, "D").Value
            Call RenderRuleSample(currentRule, objSheet.Cells(iRow, "E"))
        End If
    Next
    bookRules.Close
    
    '-- Format
    objSheet.Columns("A:A").ColumnWidth = 2
    objSheet.Columns("B:B").ColumnWidth = 12
    objSheet.Columns("C:C").ColumnWidth = 32
    objSheet.Columns("D:D").ColumnWidth = 48
    objSheet.Columns("E:E").ColumnWidth = 2
    objSheet.Cells.EntireRow.AutoFit
    
    '-- Return
    objSheet.Cells(1, 1).Select
    Set BuildSheetRules = objSheet
End Function

Private Sub BuildSheetTables(objWorkbook As Workbook, _
                                    dbType As String)
                                    
    Dim bookData As Workbook
    Set bookData = Application.Workbooks.Open(ThisWorkbook.Path & TableDataFileLocation)
    Dim sheetData As Worksheet
    Set sheetData = bookData.Worksheets(dbType)

    Dim objSheet As Worksheet
    Dim iRow As Integer
    iRow = 1
    
    While basBuildDataFileWorkbook.ImportHasTable(sheetData, iRow)
        Set objSheet = BuildSheetTable(objWorkbook)
        basBuildDataFileWorkbook.FillTableDefinitionFromDataSheet sheetData, objSheet, iRow
    Wend
    bookData.Close
                                    
End Sub

Private Function BuildSheetTable(objWorkbook As Workbook) As Worksheet
    Dim objSheet As Worksheet
    Set objSheet = objWorkbook.Sheets.Add(, objWorkbook.Sheets(objWorkbook.Sheets.Count))
    
    Call SetStyleSheet(objSheet)
    
    Dim objRange As Range
    '-- Render Table Section
    objSheet.Cells(basAppSetting.Table_Sheet_Row_TableName, "A").Value = "Table Name"
    objSheet.Cells(basAppSetting.Table_Sheet_Row_TableComment, "A").Value = "Comment"
    objSheet.Cells(basAppSetting.Table_Sheet_Row_PrimaryKey, "A").Value = "Primary Key"
    objSheet.Cells(basAppSetting.Table_Sheet_Row_ForeignKey, "A").Value = "Foreign Key"
    objSheet.Cells(basAppSetting.Table_Sheet_Row_Index, "A").Value = "Index"
    RenderRangeToLable objSheet.Range("A1:A5")
    RenderRangeMerge objSheet.Range(basAppSetting.Table_Sheet_Col_TableName & "1:H1")
    RenderRangeMerge objSheet.Range(basAppSetting.Table_Sheet_Col_TableName & "2:F2")
    RenderRangeMerge objSheet.Range(basAppSetting.Table_Sheet_Col_TableName & "3:F3")
    RenderRangeMerge objSheet.Range(basAppSetting.Table_Sheet_Col_TableName & "4:F4")
    RenderRangeMerge objSheet.Range(basAppSetting.Table_Sheet_Col_TableName & "5:F5")
    
    Set objRange = objSheet.Cells(2, Table_Sheet_Col_TableStatus)
    objRange.Value = "Status"
    RenderRangeToLable objRange
    
    Set objRange = objSheet.Cells(2, Table_Sheet_Col_Unique)
    objRange.Value = "U"
    RenderRangeToLable objRange
    
    Set objRange = objSheet.Cells(2, Table_Sheet_Col_Clustered)
    objRange.Value = "C"
    RenderRangeToLable objRange
    
    Dim objCell As Range
    Set objCell = objSheet.Cells(1, Table_Sheet_Col_TableStatus)
    objCell.Value = "<<"
    objCell.Hyperlinks.Delete
    Call objSheet.Hyperlinks.Add(objCell, _
            "", _
            "'" & Sheet_Index_Name & "'!A1")
    RenderRangeLink objCell
    RenderRangeToLable objCell
    
    RenderTableBorder objSheet.Range("A1:" & Table_Sheet_Col_TableStatus & "5")
   
    '-- Render Field Section
    Dim iFieldHeaderRow As Integer
    iFieldHeaderRow = Table_Sheet_Row_First_Column - 1
    
    Set objRange = objSheet.Cells(iFieldHeaderRow, Table_Sheet_Col_ColumnLabel)
    objRange.Value = "Column Label"
    RenderRangeToTableHeader objRange
    
    Set objRange = objSheet.Cells(iFieldHeaderRow, Table_Sheet_Col_ColumnName)
    objRange.Value = "Column Name"
    RenderRangeToTableHeader objRange
    
    Set objRange = objSheet.Cells(iFieldHeaderRow, Table_Sheet_Col_ColumnDataType)
    objRange.Value = "Type"
    RenderRangeToTableHeader objRange
    
    Set objRange = objSheet.Cells(iFieldHeaderRow, Table_Sheet_Col_ColumnNullable)
    objRange.Value = "Null"
    RenderRangeToTableHeader objRange
    
    Set objRange = objSheet.Cells(iFieldHeaderRow, Table_Sheet_Col_ColumnDefault)
    objRange.Value = "Default"
    RenderRangeToTableHeader objRange
    
    Set objRange = objSheet.Cells(iFieldHeaderRow, Table_Sheet_Col_ColumnComment)
    objRange.Value = "Comment"
    Set objRange = objSheet.Range(Table_Sheet_Col_ColumnComment & iFieldHeaderRow & ":I" & iFieldHeaderRow)
    RenderRangeMerge objRange
    RenderRangeToTableHeader objRange

    Dim iRow As Integer
    For iRow = Table_Sheet_Row_First_Column To Table_Sheet_Row_First_Column + 2
        If iRow = Table_Sheet_Row_First_Column Then
            RenderRangeMerge objSheet.Range("F" & iRow & ":I" & iRow)
        Else
            RenderRangeMerge objSheet.Range("F" & iRow & ":I" & iRow)
            '-- Slow performance
            'RenderPasteFormat objSheet.Range("A" & Table_Sheet_Row_First_Column & ":I" & Table_Sheet_Row_First_Column), _
            '    objSheet.Range("A" & iRow & ":I" & iRow)
        End If
    Next
    RenderTableBorder objSheet.Range("A" & iFieldHeaderRow & ":I" & (iFieldHeaderRow + 3))
    
    '-- set freeze panes
    objSheet.Activate
    objSheet.Range(Table_Sheet_Col_ColumnName & Table_Sheet_Row_First_Column).Select
    ActiveWindow.FreezePanes = True
    
    '-- Set column width
    objSheet.Columns("A:A").ColumnWidth = 24
    objSheet.Columns("B:B").ColumnWidth = 24
    objSheet.Columns("C:C").ColumnWidth = 20
    objSheet.Columns("D:D").ColumnWidth = 4
    objSheet.Columns("E:E").ColumnWidth = 12
    objSheet.Columns("F:F").ColumnWidth = 24
    objSheet.Columns("G:G").ColumnWidth = 2
    objSheet.Columns("H:H").ColumnWidth = 2
    objSheet.Columns("I:I").ColumnWidth = 12
    'objSheet.Cells.EntireRow.AutoFit
    
    '-- Return
    'objSheet.Cells(1, 1).Select
    Set BuildSheetTable = objSheet
End Function

Private Sub RenderRangeBold(Range As Range)
    Range.Font.Bold = True
End Sub

Private Sub RenderRangeCenter(Range As Range)
    Range.HorizontalAlignment = xlCenter
End Sub

Private Sub RenderRangeLink(Range As Range)
    Range.Font.Underline = xlUnderlineStyleNone
    Range.Style = "Normal"
End Sub

Private Sub RenderRangeGray(Range As Range)
    With Range.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Private Sub RenderRangeToLable(Range As Range)
    RenderRangeBold Range
    RenderRangeGray Range
End Sub

Private Sub RenderRangeToTableHeader(Range As Range)
    RenderRangeBold Range
    RenderRangeCenter Range
    RenderRangeGray Range
End Sub

Private Sub SetStyleSheet(objSheet As Worksheet)
    With objSheet.Cells
        .VerticalAlignment = xlTop
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Private Sub RenderRangeWrapText(Range As Range)
    With Range
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
End Sub

Private Sub RenderPasteFormat(sourceRange As Range, targetRange As Range)
    Application.ScreenUpdating = False
    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End Sub

Private Sub RenderHeader(objWorksheet As Worksheet, _
            Range As String)
            
    objWorksheet.Range(Range).Select
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub RenderRuleSample(ruleName As String, _
                Range As Range)
    If VBA.StrComp(ruleName, "Primary Key") = 0 _
        Or VBA.StrComp(ruleName, "Foreign Key") = 0 _
        Or VBA.StrComp(ruleName, "Index") = 0 _
        Or VBA.StrComp(ruleName, "Cluster") = 0 _
        Or VBA.StrComp(ruleName, "Unique") = 0 Then
        RenderTableBorderThin Range
    End If
                
End Sub

Private Sub RenderTableHeader(objWorksheet As Worksheet, _
            Range As String)

    objWorksheet.Range(Range).Select
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Private Sub RenderTableBorder(Range As Range, _
            Optional Weight As Variant = xlHairline)

    Range.Borders(xlDiagonalDown).LineStyle = xlNone
    Range.Borders(xlDiagonalUp).LineStyle = xlNone
    With Range.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = Weight
    End With
    With Range.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = Weight
    End With
    With Range.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = Weight
    End With
    With Range.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = Weight
    End With
    With Range.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = Weight
    End With
    With Range.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = Weight
    End With
End Sub

Private Sub RenderTableBorderThin(Range As Range)
    RenderTableBorder Range, xlThin
End Sub

Private Sub SetPageSetup(objWorkbook As Workbook)
    Dim objSheet As Worksheet
    
    For Each objSheet In objWorkbook.Sheets
        With objSheet.PageSetup
            .LeftHeader = "&F"
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = "&A"
            .CenterFooter = ""
            .RightFooter = "&P / &N"
            
            .PaperSize = xlPaperA4
            .Orientation = xlLandscape
        End With
    Next
End Sub

