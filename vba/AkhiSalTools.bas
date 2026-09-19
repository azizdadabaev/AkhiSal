Attribute VB_Name = "AkhiSalTools"
'==============================================================================
' AkhiSalTools - zero-install macro toolkit for Excel
'
' Import:  Alt+F11  ->  File > Import File...  ->  pick this .bas
' Run:     Alt+F8   ->  pick a macro  ->  Run
'
' Nothing here needs an add-in, an installer, or admin rights. VBA ships
' inside Excel itself.
'
' Every macro that changes the sheet writes to the ACTIVE selection or the
' active sheet's used range, and none of them can be undone by Ctrl+Z
' (VBA edits clear the undo stack), so save first. MacroBackupSheet below
' makes that one keystroke.
'==============================================================================
Option Explicit

'--- Shared helpers ----------------------------------------------------------

' The range a macro should act on: the selection when the user picked more
' than one cell, otherwise the whole used range of the sheet.
Private Function TargetRange() As Range
    If TypeName(Selection) <> "Range" Then
        Set TargetRange = Nothing
        Exit Function
    End If
    If Selection.Cells.CountLarge > 1 Then
        Set TargetRange = Intersect(Selection, ActiveSheet.UsedRange)
    Else
        Set TargetRange = ActiveSheet.UsedRange
    End If
End Function

Private Sub SpeedOn()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub SpeedOff()
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'--- 1. Backup ---------------------------------------------------------------

' Copies the active sheet next to itself before you run anything destructive.
Public Sub MacroBackupSheet()
    Dim src As Worksheet, nm As String, i As Long
    Set src = ActiveSheet
    nm = Left$(src.Name, 21) & "_bak"
    i = 1
    Do While SheetExists(nm)
        i = i + 1
        nm = Left$(src.Name, 19) & "_bak" & i
    Loop
    src.Copy After:=src
    ActiveSheet.Name = nm
    src.Activate
    MsgBox "Backup created: " & nm, vbInformation, "AkhiSal"
End Sub

Private Function SheetExists(nm As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If StrComp(ws.Name, nm, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function

'--- 2. Clean up pasted / imported data --------------------------------------

' Trims stray spaces, kills non-breaking spaces (the usual culprit when a
' VLOOKUP "should" match but doesn't), and converts numbers-stored-as-text
' back into real numbers.
Public Sub CleanImport()
    Dim rng As Range, c As Range, s As String, fixedText As Long, fixedNum As Long
    Set rng = TargetRange()
    If rng Is Nothing Then Exit Sub

    SpeedOn
    For Each c In rng.Cells
        If Not IsEmpty(c.Value) And Not c.HasFormula Then
            If VarType(c.Value) = vbString Then
                s = c.Value
                s = Replace(s, ChrW(160), " ")   ' non-breaking space
                s = Replace(s, ChrW(8206), "")   ' left-to-right mark
                s = Replace(s, ChrW(8207), "")   ' right-to-left mark
                s = Application.WorksheetFunction.Trim(s)
                If s <> c.Value Then
                    c.Value = s
                    fixedText = fixedText + 1
                End If
                If Len(s) > 0 And IsNumeric(s) Then
                    c.Value = CDbl(s)
                    fixedNum = fixedNum + 1
                End If
            End If
        End If
    Next c
    SpeedOff

    MsgBox "Cleaned " & fixedText & " text cell(s)." & vbCrLf & _
           "Converted " & fixedNum & " text value(s) to numbers.", _
           vbInformation, "AkhiSal"
End Sub

'--- 3. Phone numbers --------------------------------------------------------

' Rewrites the selection into the CRM's phone format: 998XXXXXXXXX.
' Anything it cannot read confidently is left alone and coloured, so you can
' see what needs a human.
Public Sub NormalizePhones()
    Dim rng As Range, c As Range, res As String, ok As Long, bad As Long
    Set rng = TargetRange()
    If rng Is Nothing Then Exit Sub

    SpeedOn
    For Each c In rng.Cells
        If Not IsEmpty(c.Value) And Not c.HasFormula Then
            res = PhoneUZ(CStr(c.Value))
            If Len(res) = 12 Then
                c.NumberFormat = "@"
                c.Value = res
                ok = ok + 1
            Else
                c.Interior.Color = RGB(255, 235, 156)
                bad = bad + 1
            End If
        End If
    Next c
    SpeedOff

    MsgBox "Normalized " & ok & " number(s)." & vbCrLf & _
           bad & " could not be read and are highlighted.", _
           vbInformation, "AkhiSal"
End Sub

' 998901234567 from "+998 90 123-45-67", "90 123 45 67", "8 90 1234567", ...
' Returns "" when the digits do not make a valid Uzbek mobile number.
Public Function PhoneUZ(ByVal raw As String) As String
    Dim i As Long, ch As String, d As String
    For i = 1 To Len(raw)
        ch = Mid$(raw, i, 1)
        If ch >= "0" And ch <= "9" Then d = d & ch
    Next i

    If Left$(d, 2) = "00" Then d = Mid$(d, 3)
    If Len(d) = 12 And Left$(d, 3) = "998" Then
        PhoneUZ = d
    ElseIf Len(d) = 9 Then
        PhoneUZ = "998" & d
    ElseIf Len(d) = 10 And Left$(d, 1) = "8" Then
        PhoneUZ = "998" & Mid$(d, 2)
    Else
        PhoneUZ = ""
    End If
End Function

'--- 4. Money ----------------------------------------------------------------

' UZS with space thousand separators, no decimals: 13 500 000
Public Sub FormatUZS()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Selection.NumberFormat = "# ##0"" UZS"";-# ##0"" UZS"";""-"""
End Sub

' Same, but bare numbers - better when the column feeds other formulas.
Public Sub FormatThousands()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Selection.NumberFormat = "# ##0;-# ##0;""-"""
End Sub

'--- 5. Duplicates -----------------------------------------------------------

' Highlights repeated values in the selected column(s) and reports the count.
Public Sub HighlightDuplicates()
    Dim rng As Range, c As Range, seen As Object, k As String, dupes As Long
    Set rng = TargetRange()
    If rng Is Nothing Then Exit Sub
    Set seen = CreateObject("Scripting.Dictionary")

    SpeedOn
    For Each c In rng.Cells
        If Not IsEmpty(c.Value) Then
            k = UCase$(Trim$(CStr(c.Value)))
            If seen.Exists(k) Then
                c.Interior.Color = RGB(255, 199, 206)
                seen(k).Interior.Color = RGB(255, 199, 206)
                dupes = dupes + 1
            Else
                Set seen(k) = c
            End If
        End If
    Next c
    SpeedOff

    MsgBox dupes & " duplicate cell(s) highlighted.", vbInformation, "AkhiSal"
End Sub

'--- 6. Split a table into one sheet per key ---------------------------------

' Asks which column to split on, then fans the rows out into one sheet per
' distinct value - e.g. one sheet per region, per status, or per driver.
' Assumes row 1 of the used range is the header row.
Public Sub SplitByColumn()
    Dim src As Worksheet, ur As Range, colNo As Variant
    Dim lastRow As Long, lastCol As Long, r As Long
    Dim key As String, keys As Object, ws As Worksheet

    Set src = ActiveSheet
    Set ur = src.UsedRange
    lastRow = ur.Row + ur.Rows.Count - 1
    lastCol = ur.Column + ur.Columns.Count - 1

    If lastRow - ur.Row < 1 Then
        MsgBox "Need a header row plus at least one data row.", vbExclamation, "AkhiSal"
        Exit Sub
    End If

    colNo = Application.InputBox( _
        "Split on which column number?" & vbCrLf & _
        "(1 = " & src.Cells(ur.Row, ur.Column).Text & ")", _
        "AkhiSal - Split by column", 1, Type:=1)
    If VarType(colNo) = vbBoolean Then Exit Sub
    colNo = ur.Column + CLng(colNo) - 1
    If colNo < ur.Column Or colNo > lastCol Then
        MsgBox "That column is outside the data.", vbExclamation, "AkhiSal"
        Exit Sub
    End If

    Set keys = CreateObject("Scripting.Dictionary")
    SpeedOn
    For r = ur.Row + 1 To lastRow
        key = Trim$(CStr(src.Cells(r, colNo).Text))
        If Len(key) = 0 Then key = "(blank)"
        If Not keys.Exists(key) Then
            Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
            ws.Name = SafeSheetName(key)
            src.Range(src.Cells(ur.Row, ur.Column), src.Cells(ur.Row, lastCol)).Copy _
                Destination:=ws.Cells(1, 1)
            keys.Add key, ws.Name
        End If
        Set ws = ActiveWorkbook.Worksheets(keys(key))
        src.Range(src.Cells(r, ur.Column), src.Cells(r, lastCol)).Copy _
            Destination:=ws.Cells(ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1, 1)
    Next r
    Application.CutCopyMode = False
    src.Activate
    SpeedOff

    MsgBox "Split into " & keys.Count & " sheet(s).", vbInformation, "AkhiSal"
End Sub

' Excel sheet names: 31 chars, no : \ / ? * [ ] and must be unique.
Private Function SafeSheetName(ByVal s As String) As String
    Dim bad As Variant, b As Variant, base As String, i As Long
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    For Each b In bad
        s = Replace(s, b, "-")
    Next b
    s = Trim$(s)
    If Len(s) = 0 Then s = "sheet"
    base = Left$(s, 31)
    s = base
    i = 1
    Do While SheetExists(s)
        i = i + 1
        s = Left$(base, 31 - Len(CStr(i)) - 1) & "_" & i
    Loop
    SafeSheetName = s
End Function

'--- 7. UTF-8 CSV export (keeps Cyrillic intact) -----------------------------

' Excel's own "Save as CSV" mangles Cyrillic on many Windows locales. This
' writes real UTF-8 with a BOM, so Наманган stays Наманган everywhere.
Public Sub ExportSheetToCSV_UTF8()
    Dim ws As Worksheet, ur As Range, r As Long, c As Long
    Dim lastRow As Long, lastCol As Long
    Dim line As String, out As String, cell As String
    Dim path As Variant, stream As Object

    Set ws = ActiveSheet
    Set ur = ws.UsedRange
    lastRow = ur.Row + ur.Rows.Count - 1
    lastCol = ur.Column + ur.Columns.Count - 1

    path = Application.GetSaveAsFilename( _
        InitialFileName:=ws.Name & ".csv", _
        FileFilter:="CSV file (*.csv), *.csv", _
        Title:="AkhiSal - export as UTF-8 CSV")
    If VarType(path) = vbBoolean Then Exit Sub

    SpeedOn
    For r = ur.Row To lastRow
        line = ""
        For c = ur.Column To lastCol
            cell = ws.Cells(r, c).Text
            If InStr(cell, """") > 0 Or InStr(cell, ",") > 0 _
               Or InStr(cell, vbLf) > 0 Or InStr(cell, vbCr) > 0 Then
                cell = """" & Replace(cell, """", """""") & """"
            End If
            If c > ur.Column Then line = line & ","
            line = line & cell
        Next c
        out = out & line & vbCrLf
    Next r
    SpeedOff

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2              ' text
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText out
    stream.SaveToFile CStr(path), 2
    stream.Close

    MsgBox "Saved as UTF-8:" & vbCrLf & path, vbInformation, "AkhiSal"
End Sub

'--- 8. Colour-based totals --------------------------------------------------

' Excel has no built-in "sum the yellow cells". These two fill that gap.
' Note: changing a cell's colour does not trigger recalculation - press
' Ctrl+Alt+F9 after recolouring to refresh them.
Public Function SUMBYCOLOR(sampleCell As Range, sumRange As Range) As Double
    Dim c As Range, target As Long, total As Double
    Application.Volatile
    target = sampleCell.Cells(1, 1).Interior.Color
    For Each c In sumRange.Cells
        If c.Interior.Color = target Then
            If IsNumeric(c.Value) And Not IsEmpty(c.Value) Then total = total + c.Value
        End If
    Next c
    SUMBYCOLOR = total
End Function

Public Function COUNTBYCOLOR(sampleCell As Range, countRange As Range) As Long
    Dim c As Range, target As Long, n As Long
    Application.Volatile
    target = sampleCell.Cells(1, 1).Interior.Color
    For Each c In countRange.Cells
        If c.Interior.Color = target Then n = n + 1
    Next c
    COUNTBYCOLOR = n
End Function

'--- 9. Remove fully empty rows ----------------------------------------------

Public Sub RemoveEmptyRows()
    Dim ws As Worksheet, ur As Range, r As Long, removed As Long
    Set ws = ActiveSheet
    Set ur = ws.UsedRange

    SpeedOn
    For r = ur.Row + ur.Rows.Count - 1 To ur.Row Step -1
        If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then
            ws.Rows(r).Delete
            removed = removed + 1
        End If
    Next r
    SpeedOff

    MsgBox removed & " empty row(s) removed.", vbInformation, "AkhiSal"
End Sub

'--- 10. Make a sheet printable ----------------------------------------------

' Landscape, fit to one page wide, header row repeated on every page,
' narrow margins, page numbers in the footer. Saves the usual five minutes
' of fighting Page Setup before sending a delivery list to the workshop.
Public Sub MakePrintReady()
    Dim ws As Worksheet, ur As Range
    Set ws = ActiveSheet
    Set ur = ws.UsedRange

    ws.Cells.EntireColumn.AutoFit
    With ws.PageSetup
        .PrintArea = ur.Address
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = ws.Rows(ur.Row).Address
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)
        .CenterFooter = "&P / &N"
        .RightHeader = "&D"
    End With

    MsgBox "Print layout applied. Ctrl+P to check.", vbInformation, "AkhiSal"
End Sub
