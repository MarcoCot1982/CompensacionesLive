'+------------------------------------------------------------------+
'| Author: Marco Cot         DAS:A669714                            |
'| Program which allows to check compensations before month closing.|
'| version: 2.4 [20250714]                                          |
'+------------------------------------------------------------------+
Sub DateConversion()

Application.ScreenUpdating = False

Dim RowCounter As Long
Dim DateCounter As Long
Dim RemainingComp As Double
Dim StartTime As Date
Dim EndTime As Date
Dim TimeUsed As Double
Dim TimeFormatted As String
Dim Trigger As Double

StartTime = Now

Worksheets("VALID").Visible = xlSheetVisible
Worksheets("FESTIVOSClean").Visible = xlSheetVisible
Worksheets("COMPENSACIONESClean").Visible = xlSheetVisible

Sheets("Compens").Outline.ShowLevels RowLevels:=0, ColumnLevels:=2
Sheets("TAB").Outline.ShowLevels RowLevels:=0, ColumnLevels:=2

Sheets("Festivos").Range("A:BL").Copy
Sheets("FestivosClean").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats

Sheets("VALID").Range("A:BI").Copy
Sheets("COMPENSACIONESClean").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
 
DeleteRowsNotATOS
DeleteRowsRejected
DeleteRowsAbsence


'consolidate hours
Sheets("FestivosClean").Select
Trigger = Application.WorksheetFunction.CountA(Range("A:A"))
If Trigger = 1 Then GoTo JumpHere
RowCounter = Application.WorksheetFunction.CountA(Range("BL:BL"))
Range("BL2:BL" & RowCounter).Copy
Range("P2:P" & RowCounter).PasteSpecial xlPasteValuesAndNumberFormats
Range("AG2:AG" & RowCounter).PasteSpecial xlPasteValuesAndNumberFormats
RemoveDuplicatesBasedOnTwoColumns


JumpHere:
UpdateFestivos

DateCounter = Application.WorksheetFunction.CountA(Range("BL:BL"))


UpdateComp

Sheets("Compens").Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
Sheets("TAB").Outline.ShowLevels RowLevels:=0, ColumnLevels:=1

COMP_PENDIENTES

PlannedORbaja

'Closing steps
DeleteRowsCurrentMonthTAB
DeleteRowsCurrentMonthComp

Worksheets("VALID").Visible = xlSheetHidden
Worksheets("FESTIVOSClean").Visible = xlSheetHidden
Worksheets("COMPENSACIONESClean").Visible = xlSheetHidden

'Set back to PowerBI tab
Sheets("PowerBI").Select
Range("A1").Select

EndTime = Now
TimeUsed = EndTime - StartTime
TimeFormatted = Format(TimeUsed, "nn:ss")

'Final report on popup window
RemainingComp = Application.WorksheetFunction.CountA(Range("A:A")) - 1
MsgBox "Updated at " & Now & vbCrLf _
    & vbCrLf _
    & RemainingComp & " total compensations remaining." & vbCrLf _
    & vbCrLf _
    & TimeFormatted & " total process time" _
    , vbInformation, "UPDATE COMPLETED"

End Sub
'+-----------------------------------------------------------------------------------------------+
Sub DeleteRowsNotATOS()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("FESTIVOSClean")
    
    'Find the last row with data in column S
    lastRow = ws.Cells(ws.Rows.Count, "S").End(xlUp).Row
    
    'Loop through rows in reverse order to avoid issues when deleting rows
    For i = lastRow To 2 Step -1
        If ws.Cells(i, "S").Value <> "ATOS" Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub
'+-----------------------------------------------------------------------------------------------+
Sub DeleteRowsRejected()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("FESTIVOSClean")
 
    Range("T:T").Replace "#N/A", "APPROVAL REJECTED"
    
    'Find the last row with data in column S
    lastRow = ws.Cells(ws.Rows.Count, "T").End(xlUp).Row
    
    'Loop through rows in reverse order to avoid issues when deleting rows
    For i = lastRow To 2 Step -1
        If ws.Cells(i, "T").Value = "APPROVAL REJECTED" Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub
'+-----------------------------------------------------------------------------------------------+
Sub DeleteRowsAbsence()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("FESTIVOSClean")
    
    UnionHours

    Range("AB:AB").Replace "#N/A", "Absence"

    
    'Find the last row with data in column S
    lastRow = ws.Cells(ws.Rows.Count, "AB").End(xlUp).Row
    
    'Loop through rows in reverse order to avoid issues when deleting rows
    For i = lastRow To 2 Step -1
        If ws.Cells(i, "AB").Value = "Absence" Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

'+-----------------------------------------------------------------------------------------------+
Sub RemoveDuplicatesBasedOnTwoColumns()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dict As Object

    Set ws = ThisWorkbook.Sheets("FESTIVOSClean")
    lastRow = ws.Cells(ws.Rows.Count, "BK").End(xlUp).Row
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Loop from bottom to top to safely delete rows
    For i = lastRow To 2 Step -1 ' Assuming row 1 is header
        Dim val As Variant
        val = ws.Cells(i, "BK").Value

        If dict.exists(val) Then
            ws.Rows(i).Delete
        Else
            dict.Add val, 1
        End If
    Next i

End Sub
'+-----------------------------------------------------------------------------------------------+
Sub COMP_PENDIENTES()

    Dim lastRow As Long
    Dim ws As Worksheet
    Dim tblRange As Range
    
 'FEEDS POWERBI SHEET
    
    Sheets("PowerBI").Select
    lastRow = Application.WorksheetFunction.CountA(Range("A:A"))
    Range("A2:L" & lastRow).Delete

    Sheets("TAB").Select
        
    On Error Resume Next
    ActiveSheet.ListObjects("Table4").AutoFilter.ShowAllData
    On Error GoTo 0
    
    ActiveSheet.ListObjects("Table4").Range.AutoFilter Field:=13, Criteria1:= _
        RGB(255, 192, 0), Operator:=xlFilterCellColor
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("PowerBI").Select
    Range("O2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

    lastRow = Application.WorksheetFunction.CountA(Range("O:O")) + 1

    Range("O2:O" & lastRow).Copy Destination:=Range("A2")
    Range("S2:S" & lastRow).Copy Destination:=Range("B2")
    Range("P2:P" & lastRow).Copy Destination:=Range("C2")
    Range("T2:T" & lastRow).Copy Destination:=Range("D2")
    Range("R2:R" & lastRow).Copy Destination:=Range("E2")
    Range("V2:V" & lastRow).Copy Destination:=Range("F2")
    Range("Q2:Q" & lastRow).Copy Destination:=Range("G2")
    Range("U2:U" & lastRow).Copy Destination:=Range("H2")
    Range("W2:W" & lastRow).Copy Destination:=Range("I2")
    Range("AA2:AA" & lastRow).Copy Destination:=Range("J2")
    Range("Y2:Y" & lastRow).Copy Destination:=Range("L2")
    Range("Z2:Z" & lastRow).Copy Destination:=Range("K2")
    Range("O:AA").Delete
    
    'Set the worksheet and table range
    Set ws = ThisWorkbook.Sheets("PowerBI")
    Set tblRange = ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)

    'Sort the table
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=tblRange.Columns(6), Order:=xlAscending 'DATE: Oldest to newest
        .SortFields.Add Key:=tblRange.Columns(3), Order:=xlAscending 'NAME: A to Z
        .SetRange tblRange
        .Header = xlYes
        .Apply
    End With
    
    Range("A2").Select
        
    Sheets("TAB").Select

    On Error Resume Next
    ActiveSheet.ListObjects("Table4").AutoFilter.ShowAllData
    On Error GoTo 0
    
    Sheets("PowerBI").Select
    
    AlternateRowShading
    
    Application.ScreenUpdating = True

End Sub
'+-----------------------------------------------------------------------------------------------+
Sub UpdateFestivos()

    Dim LastRowFest As Long
    Dim LastRowTAB As Long
    Dim Trigger As Double
    
    Sheets("FESTIVOSClean").Select
    Trigger = Application.WorksheetFunction.CountA(Range("A:A"))
    If Trigger = 1 Then GoTo Shortcut
    LastRowFest = Application.WorksheetFunction.CountA(Range("B:B"))
    Range("B2:B" & LastRowFest).Select
    Selection.Copy
    Sheets("TAB").Select
    LastRowTAB = Application.WorksheetFunction.CountA(Range("B:B")) + 1
    Range("A" & LastRowTAB).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("FESTIVOSClean").Select
    Range("C2:C" & LastRowFest).Select
    Selection.Copy
    Sheets("TAB").Select
    Range("B" & LastRowTAB).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("FESTIVOSClean").Select
    ActiveWindow.SmallScroll ToRight:=8
    Range("O2:O" & LastRowFest).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("TAB").Select
    Range("C" & LastRowTAB).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("FESTIVOSClean").Select
    Range("P2:P" & LastRowFest).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("TAB").Select
    Range("D" & LastRowTAB).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("FESTIVOSClean").Select
    Range("Y2:Y" & LastRowFest).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("TAB").Select
    Range("E" & LastRowTAB).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("FESTIVOSClean").Select
    Range("AA2:AA" & LastRowFest).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("TAB").Select
    Range("F" & LastRowTAB).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("FESTIVOSClean").Select
    Range("AQ2:AQ" & LastRowFest).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("TAB").Select
    Range("G" & LastRowTAB).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("FESTIVOSClean").Select
    Range("AT2:AT" & LastRowFest).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("TAB").Select
    Range("H" & LastRowTAB).Select
    ActiveSheet.Paste
    Range("Q" & LastRowTAB & ":Q" & LastRowTAB + LastRowFest - 2).Formula = "CURRENT MONTH"
    
Shortcut:
    
End Sub
'+-----------------------------------------------------------------------------------------------+
Sub DeleteRowsCurrentMonthTAB()
    Dim ws As Worksheet
    Dim FirstToDelete As Long
    Dim lastRow As Long
    Dim RowsToBeDeleted As Long
    
    Set ws = ThisWorkbook.Sheets("TAB")
    
    Sheets("TAB").Select
    lastRow = Application.WorksheetFunction.CountA(Range("B:B"))
    RowsToBeDeleted = Application.WorksheetFunction.CountA(Range("Q:Q")) - 2
    FirstToDelete = lastRow - RowsToBeDeleted
    
    If RowsToBeDeleted > -1 Then
        Range(FirstToDelete & ":" & lastRow).Delete
    End If
     
End Sub
'+-----------------------------------------------------------------------------------------------+
Sub UpdateComp()

    Dim LastRowCOMP As Long
    Dim LastRowLOG As Long
    
    Sheets("COMPENSACIONESClean").Select
    LastRowCOMP = Application.WorksheetFunction.CountA(Range("B:B"))
    Range("A2:N" & LastRowCOMP).Select
    Selection.Copy
    Sheets("Compens").Select
    LastRowLOG = Application.WorksheetFunction.CountA(Range("B:B")) + 1
    Range("A" & LastRowLOG).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Sheets("COMPENSACIONESClean").Select
    Range("O2:BD" & LastRowCOMP).Select
    Selection.Copy
    Sheets("Compens").Select
    Range("Q" & LastRowLOG).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("BM" & LastRowLOG & ":BM" & LastRowLOG + LastRowCOMP - 2).Formula = "CURRENT MONTH"

End Sub
'+-----------------------------------------------------------------------------------------------+
Sub DeleteRowsCurrentMonthComp()
    
    Dim ws As Worksheet
    Dim FirstToDelete As Long
    Dim lastRow As Long
    Dim RowsToBeDeleted As Long
    
    Set ws = ThisWorkbook.Sheets("Compens")
    
    Sheets("Compens").Select
    lastRow = Application.WorksheetFunction.CountA(Range("B:B"))
    RowsToBeDeleted = Application.WorksheetFunction.CountA(Range("BM:BM")) - 2
    FirstToDelete = lastRow - RowsToBeDeleted
    
    If RowsToBeDeleted > -1 Then
        Range(FirstToDelete & ":" & lastRow).Delete
    End If

End Sub
'+-----------------------------------------------------------------------------------------------+
Sub UnionHours()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets("FestivosClean")
    
    'Find the last row in column G with data
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    'Loop through each row in column G
    For i = 2 To lastRow
        If ws.Cells(i, "G").Value = "UNION HOURS" Then
            'If content in column G is "UNION HOURS", change content in column AB to "Productive"
            ws.Cells(i, "AB").Value = "Productive"
        End If
    Next i
End Sub
'+-----------------------------------------------------------------------------------------------+
Sub ShowPopup()

'launch window

    UserForm1.Show
End Sub
'+-----------------------------------------------------------------------------------------------+
Sub AlternateRowShading()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dataRange As Range

    'Set the worksheet
    Set ws = ThisWorkbook.Sheets("PowerBI")

    'Find the last row in column A (assuming column A always has data)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    'Define the range to apply formatting (columns A to E and G to L)
    Set dataRange = ws.Range("A2:E" & lastRow)

    'Loop through each row in the range
    For i = 2 To lastRow
        If i Mod 2 = 0 Then
            'Even row: no fill
            dataRange.Rows(i - 1).Interior.ColorIndex = xlNone
        Else
            ' Odd row: grey
            dataRange.Rows(i - 1).Interior.Color = RGB(211, 211, 211)
        End If
    Next i
    
    
     'Define the range to apply formatting (columns A to E and G to L)
    Set dataRange = ws.Range("G2:L" & lastRow)

    'Loop through each row in the range
    For i = 2 To lastRow
        If i Mod 2 = 0 Then
            'Even row: no fill
            dataRange.Rows(i - 1).Interior.ColorIndex = xlNone
        Else
            ' Odd row: grey
            dataRange.Rows(i - 1).Interior.Color = RGB(211, 211, 211)
        End If
    Next i
    
    
End Sub
'+-----------------------------------------------------------------------------------------------+
'+++CUSTOM FUNCTION++++
Function CleanDate(ByVal DateFixed As String) As String
    Dim result As String
    Dim i As Long
    Dim char As String
    Dim parts() As String
    Dim dayPart As String
    Dim monthPart As String
    Dim yearPart As String
    Dim temp As String
    
    'Step 1: Sanitize the input by retaining only digits and slashes,
    'and replacing certain delimiters with slashes
    result = ""
    For i = 1 To Len(DateFixed)
        char = Mid(DateFixed, i, 1)
        Select Case char
            Case "0" To "9", "/"
                result = result & char
            Case ",", "\", "-", "."
                result = result & "/"
            ' Ignore all other characters
        End Select
    Next i
    
    'Step 2: Split the sanitized date into components
    parts = Split(result, "/")
    
    'Ensure the date has three components (day, month, year)
    If UBound(parts) = 2 Then
        dayPart = parts(0)
        monthPart = parts(1)
        yearPart = parts(2)
        
        'Step 3: Swap day and month if the month value exceeds 12
        If IsNumeric(monthPart) And val(monthPart) > 12 Then
            ' Swap dayPart and monthPart
            temp = dayPart
            dayPart = monthPart
            monthPart = temp
        End If
        
        'Reconstruct the date string
        result = dayPart & "/" & monthPart & "/" & yearPart
    End If
    
    CleanDate = result
End Function

'+-----------------------------------------------------------------------------------------------+

'+-----------------------------------------------------------------------------------------------+
Sub PlannedORbaja()

    Dim lastRow As Long
    Dim wsSoporte As Worksheet
    Dim wsPowerBI As Worksheet

    ' Set worksheet references
    Set wsSoporte = Sheets("SOPORTE")
    Set wsPowerBI = Sheets("PowerBI")

    ' Find and delete rows in SOPORTE
    With wsSoporte
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        If lastRow >= 3 Then
            .Rows("3:" & lastRow).Delete
        End If
    End With

    ' Copy from PowerBI to SOPORTE
    With wsPowerBI
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range("A2:L" & lastRow).Copy
    End With

    wsSoporte.Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False

    ' Copy column M from SOPORTE and paste in PowerBI L2
    With wsSoporte
        .Range("M2:M" & lastRow).Copy
    End With

    wsPowerBI.Range("L2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False

End Sub
