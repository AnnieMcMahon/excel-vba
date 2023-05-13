Sub CentageFile()
'
' CentageFile Macro
' Formats TB CSV file for Centage uploads
' Edited 1/2023 - changed instructions
'
' Keyboard Shortcut: Ctrl+Shift+C
'
Dim I As Integer 'Total number of rows
Dim J As Integer 'Counter
Dim CompanyName As String 'Company name
Dim DedCode As String 'Deduction code
Dim Loc As String 'Location code
Dim MacroFile As String 'Macro file name
Dim CSVFile As String 'CSV file name
Dim MarketFile As String 'Market tab name
Dim MarketCode As String 'Market code name
Dim CompanyCode As String 'First 3 digits of file name

MacroFile = "Centage Macro File 2023-01.xlsm"
CSVFile = ActiveWorkbook.Name
CompanyCode = Left(CSVFile, 3)
If CompanyCode = "080" Then
    MarketFile = ActiveSheet.Name
    MarketCode = Left(MarketFile, 5)
End If

'Find how many rows the report has and remove 0.00 values
    I = 1
    Do While Range("A" & I) <> ""
        If Range("B" & I) = 0 Then
            Rows(I & ":" & I).Select
            Selection.Delete Shift:=xlUp
        Else
            I = I + 1
        End If
    Loop
        
'Copy necessary data to macro file
Range("A1:B" & I).Copy
Windows(MacroFile).Activate
Sheets("Macro").Select
Range("A1").Select
ActiveSheet.Paste

'Close CSV file without saving
If CompanyCode <> "080" Then
    Windows(CSVFile).Activate
    ActiveWindow.Close
    Windows(MacroFile).Activate
End If

'Insert company name
Sheets("Company Names").Select
J = 1
If CompanyCode <> "080" Then
    Do While CompanyCode <> Range("A" & J)
        J = J + 1
    Loop
    CompanyName = Range("B" & J).Value
Else
    Do While MarketCode <> Range("D" & J)
        J = J + 1
    Loop
    CompanyName = Range("E" & J).Value
End If

Sheets("Macro").Select
Columns("A:A").Select
Selection.Insert Shift:=xlToRight
J = 1
Do While J < I
    Range("A" & J) = CompanyName
    J = J + 1
Loop


'Move amounts in Debit and Credit columns
    J = 1
    Do While J < I
        If Range("C" & J) < 0 Then
            Range("H" & J) = Range("C" & J) * -1
            Range("G" & J) = 0
        Else
            Range("G" & J) = Range("C" & J)
            Range("H" & J) = 0
        End If
        Range("C" & J).ClearContents
        J = J + 1
    Loop
        
'Add entity name to GL code (certain companies)
    Select Case CompanyName
        Case Is = "True World Marine"
            Loc = "Marine-"
        Case Is = "True World Outdoor"
            Loc = "TWOUT-"
        Case Is = "Flying Ocean"
            Loc = "FO-"
        Case Is = "TWF-UK"
            Loc = "TWF-UK-"
        Case Is = "TWF-Spain"
            Loc = "TWS-"
        Case Is = "TW Korea"
            Loc = "TWK-"
        Case Is = "TW Japan"
            Loc = "TWJ-"
        Case Is = "TWF LTD"
            Loc = "TWFIRE-"
    End Select
    J = 1
    Do While J < I
        DedCode = Loc & Range("B" & J)
        Range("B" & J) = DedCode
        J = J + 1
    Loop
    
'Format columns
    Columns("A:H").Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Bold = False
        .Italic = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Columns("G:H").Select
    Selection.NumberFormat = "0.00"
    Columns("A:H").EntireColumn.AutoFit
    Range("A1:H" & I - 1).Select
End Sub
