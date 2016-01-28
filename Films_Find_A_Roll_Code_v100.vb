Sub DataPull()
    
'    This macro iniaites the pull of SAP information from the following sources in order to find aged rolls to reslit into CRIMs rolls.
'    Batch Info for FG in 1197 (LQUA)
'    Age of those batches (ZMCH1)
'    Material Descriptions(ZMARAC)
'    Material Safety Stocks (ZMARACVKE)
'    Demand for Materials in 1197 (ZM44)
    
'    This part of the macro determines the filepath of this file and looks for an initial VBS file in the same folder.
    Dim filepath As String
    Dim explorerpath As String
    filepath = Application.ActiveWorkbook.Path
    explorerpath = "explorer.exe " & filepath & "\AgedFilms_1_InitialRun.vbs"
    Shell explorerpath, vbNormalFocus
    
End Sub

Sub ImportAged()
    
'    This macro imports the listing of finished goods in 1197 from LQUA.
'    This is the first step in determing which finished goods are aged.
'    AgedFilms_1_InitialRun.vbs calls this subroutine to run after the VBS script finishes.

'    In this section we clear out whatever results were previously in the file.
    Sheets("AgedFilms").Select
    Cells.AutoFilter
    Cells.Select
    Cells.Clear
    Range("A1").Select
    
'    This section determines the path of the current workbook and using that path
'    attempts to import the SAP output which is in text format.
    Dim programPath As String
    programPath = Application.ActiveWorkbook.Path
    Dim scriptPath As String
    scriptPath = programPath & "\FGListing.txt"

    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & scriptPath, Destination:= _
        Range("$A$1"))
        .Name = "FGListing_2"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 4
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With

'    The section formats the SAP import from LQUA.
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Rows("1:1").Select
    Selection.Font.Bold = True
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    
'    The below line sends us to the next subroutine.
    Application.Run "CopyBatches"

End Sub

Sub CopyBatches()

'    This copies batches over so they can be dropped into SAP, transaction ZMCH1.
'    The purpose of this macro and script are to pull in the creation date of the batches.

    Sheets("AgedFilms").Select
    Dim lastRow As Long
    lastRow = Sheets("AgedFilms").Cells(Rows.Count, "C").End(xlUp).Row
    Range("C2:C" & lastRow).Select
    Selection.Copy
     
    Dim filepath As String
    Dim explorerpath As String
    filepath = Application.ActiveWorkbook.Path
    explorerpath = "explorer.exe " & filepath & "\AgedFilms_2_BatchDates.vbs"
    Shell explorerpath, vbNormalFocus
     
End Sub

Sub ImportBatchDates()

'    This macro imports the batch dates that were output from SAP.

    Sheets("BatchDates").Select
    Range("A1").Select
    Sheets("BatchDates").Cells.Clear
    
    Dim identifierPath2 As String
    identifierPath2 = Left(Application.CommandBars("Web").Controls("Address:").Text, Len(Application.CommandBars("Web").Controls("Address:").Text) - (Len("Films_Find_A_Roll.xls")))
    
    With ActiveSheet.QueryTables.Add(Connection:= _
      "URL;" + identifierPath2 & "\" & "ExportBatchDates.htm", _
         Destination:=Range("a1"))
    
      .BackgroundQuery = True
      .TablesOnlyFromHTML = True
      .Refresh BackgroundQuery:=False
      .SaveData = True
    End With
    
    Application.Run "VLOOKUPDATES"

End Sub

Sub VLOOKUPDATES()

'    This macro addes the imported batch dates to the original material rows output from LQUA.

    Sheets("AgedFilms").Select
    
    Dim lastRow As Long
    lastRow = Sheets("AgedFilms").Cells(Rows.Count, "A").End(xlUp).Row

    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Date"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],BatchDates!C[-7]:C[-6],2,FALSE)"
    Range("H2").Select
    Selection.AutoFill Destination:=Range("H2:H" & lastRow)
    Columns("H:H").Select
    Selection.NumberFormat = "mm/dd/yy;@"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=TODAY()-RC[-1]"
    Range("I2").Select
    Selection.AutoFill Destination:=Range("I2:I" & lastRow)
    Columns("I:I").Select
    Selection.NumberFormat = "0"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Days Old"
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("H:I").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select
    
    ActiveSheet.Range("$A$1:$I$" & lastRow).AutoFilter Field:=9, Criteria1:="=#N/A"
    Range("A1").Select
    
    ActiveCell.Offset(1, 0).Select

    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.EntireRow.Delete
    Columns("A:I").Select
    Selection.AutoFilter
    
    Application.Run "CopyMats"
    
End Sub

Sub CopyMats()

'    This macro copies the material numbers so that material profiles (ZMARAC),
'    material SS (ZMARACVKE), and material demand (ZM44) can all be pulled in and merged.

    Sheets("AgedFilms").Select
    Dim lastRow As Long
    lastRow = Sheets("AgedFilms").Cells(Rows.Count, "C").End(xlUp).Row
    Range("A2:A" & lastRow).Select
    Selection.Copy
       
    Dim filepath As String
    Dim explorerpath As String
    filepath = Application.ActiveWorkbook.Path
    explorerpath = "explorer.exe " & filepath & "\AgedFilms_3_MatProfiles.vbs"
    Shell explorerpath, vbNormalFocus
        
End Sub

Sub ImportDemand()

'    This macro imports the ZM44 output and formats it so we can check the aged materials against demand.
'    The purpose of this macro is to provide information to later determine if we should not reslit and item
'    due to the item having open demand.

    Sheets("DemandCheck").Select
    Range("A1").Select
    Sheets("DemandCheck").Cells.Clear
    
    Dim programPath As String
    programPath = Application.ActiveWorkbook.Path
    Dim scriptPath As String
    scriptPath = programPath & "\filmsdemandcheck.txt"
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & scriptPath, _
        Destination:=Range("$A$1"))
        .Name = "FilmsMatProfiles_13"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    Columns("A:B").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:D").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    Rows("1:8").Select
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Rows("1:1").Select
    Selection.Font.Bold = True
    Dim LR As Long, i As Long
    LR = Range("I" & Rows.Count).End(xlUp).Row
    For i = LR To 2 Step -1
    If Range("I" & i).Value = 0 Then Rows(i).Delete
        Next i
    On Error Resume Next
    Columns("A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select

    Application.Run "ImportSS"

End Sub

Sub ImportSS()

'    This macro imports the ZMARACVKE output and formats it so we can check if aged materials have a SS.
'    The purpose of this macro is to provide information to later determine if we should not reslit and item
'    due to the item already being a CRIMS material.

    Sheets("SSCheck").Select
    Range("A1").Select
    Sheets("SSCheck").Cells.Clear
    
    Dim programPath As String
    programPath = Application.ActiveWorkbook.Path
    Dim scriptPath As String
    scriptPath = programPath & "\FilmsSSCheck.txt"

    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & scriptPath, _
        Destination:=Range("$A$1"))
        .Name = "FilmsMatProfiles_13"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Rows("3:3").Select
    Selection.Delete Shift:=xlUp
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Selection.Font.Bold = True
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    
    Dim lastRow As Long
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    Columns("A:D").Select
    ActiveSheet.Range("$A$1:$I$" & lastRow).RemoveDuplicates Columns:=Array(1, 2, 3, 4), _
        Header:=xlYes
    Range("A1").Select

    Application.Run "ImportMatProfiles"

End Sub

Sub ImportMatProfiles()

'    This macro imports the ZMARAC output and formats it so material descriptions can be matched with batches.
'    The purpose of this macro is to provide information to determine possible reslit opportunities.

    Sheets("MaterialProfiles").Select
    Range("A1").Select

    Sheets("MaterialProfiles").Cells.Clear
    
    Dim programPath As String
    programPath = Application.ActiveWorkbook.Path
    Dim scriptPath As String
    scriptPath = programPath & "\FilmsMatProfiles.txt"
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & scriptPath, _
        Destination:=Range("$A$1"))
        .Name = "FilmsMatProfiles_13"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Rows("3:3").Select
    Selection.Delete Shift:=xlUp
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Selection.Font.Bold = True
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select

    Dim lastRow As Long
    lastRow = Range("A" & Rows.Count).End(xlUp).Row

    Columns("A:D").Select
    ActiveSheet.Range("$A$1:$D$" & lastRow).RemoveDuplicates Columns:=Array(1, 2, 3, 4), _
        Header:=xlYes
    Range("A1").Select

'    This section adds profit center and material description to the aged films list.
    Sheets("AgedFilms").Select
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Profit Center"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Material Description"
    Range("A1:K1").Select
    Selection.AutoFilter
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("J2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-9],MaterialProfiles!C[-9]:C[-7],3,FALSE)"
    Range("J2").Select
    Selection.AutoFill Destination:=Range("J2:J" & lastRow)
    Range("J2:J" & lastRow).Select
    Range("K2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-10],MaterialProfiles!C[-10]:C[-7],4,FALSE)"
    Range("K2").Select
    Selection.AutoFill Destination:=Range("K2:K" & lastRow)
    Range("K2:K" & lastRow).Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("J:K").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J3").Select
    Application.CutCopyMode = False
    Range("A1").Select

    Application.Run "AddFormulas"

End Sub

Sub AddFormulas()

'    This macro adds in all the formulas and values from the formulas tab.  If you need to the edit the formulas
'    Do so in the Formulas tab.  The exceptions below are for the SS Check and the Demand Check which should be
'    Hardcoded (Excel will edit the formulas due to how we've coded the data cleanup with row deletes)

    Sheets("Formulas").Select
    Range("AF2").Select
    ActiveCell.Formula = "=VLOOKUP(A2,SSCheck!$A:$I,9,FALSE)"
    Range("AG2").Select
    ActiveCell.Formula = "=IF(ISERROR(VLOOKUP(A2,DemandCheck!C:C,1,FALSE)),""No"",""Yes"")"
    Range("L1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("AgedFilms").Select
    
    Dim lastRow As Long
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    Range("L1").Select
    ActiveSheet.Paste
    Range("L2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("L2:AG" & lastRow)
    Range("L2:AG" & lastRow).Select
    Columns("L:AG").Select
    Columns("L:AG").EntireColumn.AutoFit
    Cells.Select
    Range("L1").Activate
    Cells.EntireColumn.AutoFit
    Columns("L:L").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select
    ActiveWorkbook.Save

    Sheets("OrdersToAction").Select
    Cells.Select
    Selection.Delete Shift:=xlUp

    Application.Run "PrepareForExport"

End Sub
Sub PrepareForExport()

'    This marco prepares the matches for pasting into the Google Doc.
'    Matches where the CRIMS item was too large, aged item had SS, or aged item had demand are excluded.

    Sheets("AgedFilms").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit

    Dim lastRow As Long
    lastRow = Range("A" & Rows.Count).End(xlUp).Row

    ActiveSheet.Range("$A$1:$AG$1000").AutoFilter Field:=31, Criteria1:="Y"
    ActiveSheet.Range("$A$1:$AG$1000").AutoFilter Field:=32, Criteria1:="=0", _
        Operator:=xlAnd
    ActiveSheet.Range("$A$1:$AG$1000").AutoFilter Field:=33, Criteria1:="=*No*" _
        , Operator:=xlAnd
    
    Cells.Select
    Selection.Copy
    Sheets("OrdersToAction").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Take From"
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Columns("J:J").Select
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("E:E").Select
    Selection.Cut
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight

    Columns("AB:AB").Select
    Selection.Cut

    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight

    Columns("AB:AB").Select
    Selection.Cut

    Application.CutCopyMode = False

    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("I:I").Select


    Columns("I:I").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlToLeft
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "CRIM Description"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "CRIM Material"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Reclass to Material"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Sales Order / Conrel"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "LI"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Due From 1110"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Prod Order"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Completed"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "1197 STO #"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Comments"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],CRIMS!C[-4]:C[-3],2,FALSE)"
    Range("F3").Select
    Columns("F:F").EntireColumn.AutoFit
    Range("F2").Select
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    Selection.AutoFill Destination:=Range("F2:F" & lastRow)
    Range("F2:F36").Select
    Columns("F:F").EntireColumn.AutoFit
    Columns("G:I").Select
    Columns("G:I").EntireColumn.AutoFit
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=TODAY()+14"
    Range("J3").Select
    Columns("J:J").EntireColumn.AutoFit
    Range("J2").Select
    Selection.AutoFill Destination:=Range("J2:J" & lastRow)
    Range("J2:J" & lastRow).Select
    Columns("K:K").EntireColumn.AutoFit
    Range("L2").Select
    Columns("M:M").EntireColumn.AutoFit
    Range("N2").Select
    Columns("N:N").EntireColumn.AutoFit
    ActiveCell.FormulaR1C1 = "CRIM RESLIT PROGRAM"
    Range("N3").Select

    Columns("N:N").EntireColumn.AutoFit
    Range("N2").Select
    Selection.AutoFill Destination:=Range("N2:N" & lastRow)
    Range("N2:N" & lastRow).Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Rows("1:1").Select
    Selection.Font.Bold = True
    Columns("A:A").EntireColumn.AutoFit

    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Batch #"

    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    
End Sub


Sub folderPath()

'    This macro is meant for troubleshooting.  Run it if there is some kind of error with pathing to show the path.

    Dim filepath As String
    filepath = Application.ActiveWorkbook.Path
    MsgBox filepath

End Sub


Sub SaveAndClose()

    Application.CutCopyMode = False
    ActiveWorkbook.Close SaveChanges:=True

End Sub
