Sub UltiproReport()

' Ultipro Report Macro
' Macro recorded 12/17/2021 by Annie McMahon
' Revised 1/27/2022 - calculate total wages when two pay dates
' Revised 5/2022 - integrate Ultipro reports
' Revised 6/2/2022 - eliminate $0.00 Wage rows
' Revised 9/1/2022 - added INLM to NTWF group and fixed glitch when only one employee
' Revised 1/5/2023 - fixed a glitch when employees have the same exact name
' Shortcut: CTRL+SHIFT+U
'
Dim I As Integer 'Total number of rows
Dim J As Integer 'Counter
Dim E As Integer 'Employee - to transfer all data on one same row
Dim L As Integer 'Location row for summary
Dim N As Integer 'Non-TWF starting row
Dim R As Integer 'First row of the group needing a total
Dim DedCode As String 'Deduction code
Dim Loc As String 'Location code

Dim Wage As Double
Dim Contrib As Double
Dim TotalMatch As Double
Dim Loan1, Loan2 As Double

Dim TotalWage As Double
Dim Total401K, TotalRoth, GTotalMatch As Double
Dim TotalLoan1, TotalLoan2 As Double

'Delete info from Final tab
    Sheets("Final").Select
    Columns("A:N").Delete
    Sheets("Original").Select
    Columns("A:J").Delete

'Copy info to Original tab
    Sheets("Ultipro Report").Select
    Columns("A:J").Copy
    Sheets("Original").Select
    Range("A1").Select
    ActiveSheet.Paste

'Find how many rows the report has
    I = 1
    Do While Range("A" & I) <> ""
        I = I + 1
    Loop
    I = I - 1
    Columns("A:M").EntireColumn.AutoFit
    
'Copy Original report to Final tab and adjust needed columns
    Columns("A:J").Copy
    Sheets("Final").Select
    Columns("A:J").Insert
    Columns("E:E").Delete
    Columns("F:N").Insert
    
'Add and format column names
    Range("A1") = "Partic ID"
    Range("B1") = "Loc"
    Range("C1") = "Pay Date"
    Range("D1") = "Name"
    Range("E1") = "Wage"
    Range("F1") = "401K"
    Range("G1") = "Roth"
    Range("H1") = "401K CM"
    Range("I1") = "Roth CM"
    Range("J1") = "Total CM"
    Range("K1") = "Loan1"
    Range("L1") = "Loan2"
    Range("M1") = "PUCC"
    
    Rows("1:1").Select
    With Selection.Font
        .Name = "Arial"
        .Size = 9
        .Bold = True
    End With

'Align deductions on one row per employee
    J = 2
    E = 2
    I = I + 1
    Do While J < I
        DedCode = Range("O" & J)
        If Range("E" & J) = 0 Then
            Range("J" & J) = "D"
        Else
            
        Select Case DedCode
            Case "C401P", "C401D"
                If Range("F" & E) = 0 Then
                    Range("F" & E) = Range("Q" & J)
                    Range("H" & E) = Range("R" & J)
                Else
                    Range("F" & E) = Range("Q" & J) + Range("F" & E)
                    Range("H" & E) = Range("R" & J) + Range("H" & E)
                    Range("E" & E) = Range("E" & J) + Range("E" & E)
                End If
            Case "CROTP", "CROTD"
                If Range("G" & E) = 0 Then
                    Range("G" & E) = Range("Q" & J)
                    Range("I" & E) = Range("R" & J)
                Else
                    Range("G" & E) = Range("Q" & J) + Range("G" & E)
                    Range("I" & E) = Range("R" & J) + Range("I" & E)
                    Range("E" & E) = Range("E" & J) + Range("E" & E)
                End If
            Case "401KL"
                Range("K" & E) = Range("Q" & J)
            Case "401L2"
                Range("L" & E) = Range("Q" & J)
            Case "TPUCC"
                Range("M" & E) = Range("Q" & J)
                If Range("F" & E) + Range("G" & E) + Range("K" & E) + Range("L" & E) = 0 Then
                    Range("J" & J) = "D"
                Else
                    Range("M" & E) = Range("Q" & J)
                End If
        End Select
        End If
        If J <> E Then
            Range("J" & J) = "D"
        End If
        J = J + 1
        If Range("B" & J) <> Range("B" & E) Then
            E = J
        Else
            If Range("A" & J) <> Range("A" & E) Then
                E = J
            End If
        End If
    Loop
    
'Delete duplicate rows and columns no longer needed
    J = 2
    Do While Range("A" & J) <> ""
        If Range("J" & J) = "D" Then
            Range("A" & J, "R" & J).Delete
            I = I - 1
        Else
            J = J + 1
        End If
    Loop
    
    Columns("O:R").Delete
    
'Calculate total company match
    J = 2
    Do While J < I
        Range("J" & J) = Range("H" & J) + Range("I" & J)
        J = J + 1
    Loop
    
'Deduct PUCC from Gross Wage
    J = 2
    Do While J < I
        Range("E" & J) = Range("E" & J) - Range("M" & J)
        J = J + 1
    Loop
    Columns("M:N").Delete
    
'Assign location code
Columns("A:A").Insert

I = 2
Range("A1") = "Code"
Do While Range("C" & I) <> ""
    Loc = Range("C" & I)
    Select Case Loc
        Case "INLM", "SEARA", "OWMIN", "OWMLL", "SAKUR", "TWMAD", "TWG", "WWJP"
            Range("A" & I) = "NTWF"
        Case Else
            Range("A" & I) = "ATWF" 'I added "A" to sort: TWF first then others
        End Select
        I = I + 1
    Loop

'Sort by location code and name
    Range("A:M").Sort Key1:=Range("A1"), Order1:=xlAscending, _
    Key2:=Range("C1"), Order2:=xlAscending, _
    Key3:=Range("E1"), Order2:=xlAscending, _
    Header:=xlGuess
    
'*********************** PART 2: CREATE SUMMARY AND INSERT TOTALS *************************

    J = 2 'First row of the report
    L = 2 'First row of the summary
    R = 2 'First row of company group needing total at the end
    Range("P1") = "Summary"
    Range("Q1") = "Contrib"
    Range("R1") = "Loans"
    Range("P1:R1").Font.Bold = True
    
'Calculate totals for TWF
    Wage = Range("F" & J)
    Contrib = Range("G" & J)
    Roth = Range("H" & J)
    TotalMatch = Range("K" & J)
    Loan1 = Range("L" & J)
    Loan2 = Range("M" & J)
    Loc = Range("A" & J)
    J = J + 1
        
    Do While Range("A" & J) = "ATWF"
        Wage = Wage + Range("F" & J)
        Contrib = Contrib + Range("G" & J)
        Roth = Roth + Range("H" & J)
        TotalMatch = TotalMatch + Range("K" & J)
        Loan1 = Loan1 + Range("L" & J)
        Loan2 = Loan2 + Range("M" & J)
        J = J + 1
    Loop
    Range("P" & L) = "TWF"
    Range("Q" & L) = Contrib + Roth + TotalMatch
    Range("R" & L) = Loan1 + Loan2
    L = L + 1
                
'Calculate totals for NTWF
    Wage = Range("F" & J)
    Contrib = Range("G" & J)
    Roth = Range("H" & J)
    TotalMatch = Range("K" & J)
    Loan1 = Range("L" & J)
    Loan2 = Range("M" & J)
    Loc = Range("C" & J)
    J = J + 1
    
    If Range("A" & J) = "" Then
        If Loc <> "" Then
            Range("P" & L) = Loc
            Range("Q" & L) = Contrib + Roth + TotalMatch
            Range("R" & L) = Loan1 + Loan2
            L = L + 1
            J = J + 1
        End If
    End If
   
    Do While Range("A" & J) = "NTWF"
        Do While Range("C" & J) = Range("C" & J - 1)
            Wage = Wage + Range("F" & J)
            Contrib = Contrib + Range("G" & J)
            Roth = Roth + Range("H" & J)
            TotalMatch = TotalMatch + Range("K" & J)
            Loan1 = Loan1 + Range("L" & J)
            Loan2 = Loan2 + Range("M" & J)
            Loc = Range("C" & J)
            J = J + 1
        Loop
        If Loc <> "" Then
            Range("P" & L) = Loc
            Range("Q" & L) = Contrib + Roth + TotalMatch
            Range("R" & L) = Loan1 + Loan2
            L = L + 1
        End If
        
        Wage = Range("F" & J)
        Contrib = Range("G" & J)
        Roth = Range("H" & J)
        TotalMatch = Range("K" & J)
        Loan1 = Range("L" & J)
        Loan2 = Range("M" & J)
        Loc = Range("C" & J)
        J = J + 1
        If Range("A" & J) = "" Then
            If Loc <> "" Then
                Range("P" & L) = Loc
                Range("Q" & L) = Contrib + Roth + TotalMatch
                Range("R" & L) = Loan1 + Loan2
                L = L + 1
                J = J + 1
            End If
        End If
    Loop
    
    J = J - 1
    
    Range("F" & J) = WorksheetFunction.Sum(Range("F" & R & ":F" & J - 1)) 'Total Wages
    Range("G" & J) = WorksheetFunction.Sum(Range("G" & R & ":G" & J - 1)) 'Total 401K
    Range("H" & J) = WorksheetFunction.Sum(Range("H" & R & ":H" & J - 1)) 'Total Roth
    Range("K" & J) = WorksheetFunction.Sum(Range("K" & R & ":K" & J - 1)) 'Total Match
    Range("L" & J) = WorksheetFunction.Sum(Range("L" & R & ":L" & J - 1)) 'Total Loan1
    Range("M" & J) = WorksheetFunction.Sum(Range("M" & R & ":M" & J - 1)) 'Total Loan2
    Range("E" & J + 1) = "TTL Submit"
    Range("F" & J + 1) = WorksheetFunction.Sum(Range("G" & J & ":K" & J))
    Range("L" & J + 1) = "TTL Loan"
    Range("M" & J + 1) = WorksheetFunction.Sum(Range("L" & J & ":M" & J))
    Range("A" & J & ":N" & J + 1).Font.Bold = True

'Format columns and delete CM extra columns and code column
    Columns("I:J").Delete
    Columns("A:A").Delete
    Columns("K:K").Delete
    Columns("C:C").NumberFormat = "mm/dd/yyyy"
    Columns("E:R").NumberFormat = "0.00"
    Columns("A:R").EntireColumn.AutoFit
    Range("A1").Select
    
    L = L - 1

'Format summary
    Range("M2:N" & L).Select
    With Selection
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlRight
    End With
    Range("L2:L12").HorizontalAlignment = xlLeft
    Range("L2:N" & L).Font.Bold = False
    
'Copy info to Ultipro Report tab
    Columns("A:N").Select
    Selection.Copy
    Sheets("Ultipro Report").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("M2:N" & L).Select
End Sub

Sub PayrollReport()
'
' Import Ultipro report's data onto Partic tab
' Keyboard Shortcut: Ctrl+Shift+R
'
Dim I As Integer 'Total number of rows in payroll report
Dim J As Integer 'Total number of rows in Ultipro report
                  
'Find how many rows Partic tab has
    Sheets("Partic").Select
    I = 2
    Do While Range("A" & I) <> ""
        I = I + 1
    Loop
    I = I - 1
            
'Copy and paste pay date
    Sheets("Ultipro Report").Select
    Range("C2").Copy
    Sheets("Partic").Select
    Range("J2:J" & I).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
'Copy and paste vlookup formulas
    Sheets("Ultipro Report").Select
    Range("P2:V2").Copy 'Formulas
    Sheets("Partic").Select
    Range("K2:Q" & I).Select
    ActiveSheet.Paste

'Copy and paste values only
    Range("K2:Q" & I).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    
'Add totals
    Sheets("Ultipro Report").Select
    J = 2
    Do While Range("A" & J) <> ""
        J = J + 1
    Loop
    Range("E" & J & ":J" & J).Copy
    Sheets("Partic").Select
    Range("L" & I + 4).Select
    ActiveSheet.Paste

'Erase data on Vanguard tabs and Template
    Sheets("File Export").Select
    Columns("A:L").Delete
    Range("A1").Select
    
    Sheets("Print").Select
    Columns("A:H").Delete
    Range("A1").Select


    Sheets("Partic").Select
    Range("A1").Select
    
End Sub

Sub FileExport()
'
' Keyboard Shortcut: Ctrl+Shift+F
' Populates File Export tab to copy to Access database
'
Dim I As Integer 'Counter
Dim J As Integer 'Counter
Dim ReportName As String 'This workbook

ReportName = ActiveWorkbook.Name
If Left(ReportName, 9) = "_Template" Then
    MsgBox "Please save with payroll date before continuing."
Else
    
'Find how many rows the Partic report has
    Sheets("Partic").Select
    I = 2
    Do While Range("A" & I) <> ""
        I = I + 1
    Loop
    I = I - 1
    
'Copy data on File Export tab
    Sheets("Partic").Select
    Range("A1:Q" & I).Copy
    Sheets("File Export").Select
    Range("A1").Select
    ActiveSheet.Paste

'Format data
    Columns("A:A").Insert
    Columns("E:J").Delete
    Columns("F:F").Delete
    Columns("E:E").NumberFormat = "m/d/yyyy"
    Columns("D:D").NumberFormat = "0"
    Columns("F:K").NumberFormat = "0.00"

'Delete empty rows
    J = 2
    Do While Range("B" & J) <> ""
        Range("L" & J).Select
        Selection.FormulaR1C1 = "=SUM(RC[-5]:RC[-1])"
        If Range("L" & J) = 0 Then
            Rows(J & ":" & J).Delete Shift:=xlUp
        Else
            J = J + 1
        End If
    Loop
    J = J - 2
    Columns("L:L").Delete
    Columns("A:K").EntireColumn.AutoFit
    Rows("1:1").Delete
    Range("A1:K" & J).Select
    
End If
End Sub

Sub PrintPDF()
'
' Keyboard Shortcut: Ctrl+Shift+Q
' Populates Print tab to create PDF files to e-mail to each branch
'
Dim I As Integer 'Counter
Dim J As Integer 'Counter
Dim ReportName As String 'This workbook

ReportName = ActiveWorkbook.Name
If Left(ReportName, 9) = "_Template" Then
    MsgBox "Please save with payroll date before continuing."
Else
    
'Find how many rows the Partic report has
    Sheets("Partic").Select
    I = 2
    Do While Range("A" & I) <> ""
        I = I + 1
    Loop
    I = I - 1
    
'Create header for Print tab
    Sheets("Print").Select
    Range("A1") = "Loc"
    Range("B1") = "Name"
    Range("C1") = "401K"
    Range("D1") = "Roth"
    Range("E1") = "CoMatch"
    Range("F1") = "Loan1"
    Range("G1") = "Loan2"
    Rows("1:1").Font.Bold = True

'Enter formulas in first row
    Range("A2") = "=Partic!A2"
    Range("B2") = "=Partic!B2"
    Range("C2") = "=Partic!M2"
    Range("D2") = "=Partic!N2"
    Range("E2") = "=Partic!O2"
    Range("F2") = "=Partic!P2"
    Range("G2") = "=Partic!Q2"
    Range("H2") = "=SUM(C2:G2)"
    
    Range("C2:H2").NumberFormat = "0.00"

'Copy formulas all the way down
    Range("A2:H2").Select
    Selection.AutoFill Destination:=Range("A2:H" & I)
    Columns("A:I").EntireColumn.AutoFit
    
'Delete $0 values
 J = 2
 Do While Range("A" & J) <> ""
    If Range("H" & J) = 0 Then
        Rows(J & ":" & J).Delete
    Else
        J = J + 1
    End If
 Loop
       
'Insert totals for each location
    J = 2
    Do While Range("A" & J) <> ""
        K = Range("C" & J)
        R = Range("D" & J)
        M = Range("E" & J)
        L1 = Range("F" & J)
        L2 = Range("G" & J)
        J = J + 1
        Do While Range("A" & J) = Range("A" & J - 1)
            K = K + Range("C" & J)
            R = R + Range("D" & J)
            M = M + Range("E" & J)
            L1 = L1 + Range("F" & J)
            L2 = L2 + Range("G" & J)
            J = J + 1
        Loop
        Range("A" & J).Select
        Selection.EntireRow.Insert
        Selection.EntireRow.Insert
        Selection.EntireRow.Insert
        Selection.EntireRow.Insert
        Selection.EntireRow.Insert
        Selection.EntireRow.Insert
        Selection.EntireRow.Insert
        Range("B" & J) = "Total"
        Range("C" & J) = K
        Range("D" & J) = R
        Range("E" & J) = M
        Range("F" & J) = L1
        Range("G" & J) = L2
        Range("B" & J + 2) = "Total Contributions"
        Range("C" & J + 2) = K + R + M
        Range("B" & J + 3) = "Total Loans"
        Range("C" & J + 3) = L1 + L2
        Range("B" & J + 4) = "Grand Total"
        Range("C" & J + 4) = K + R + M + L1 + L2
        Range("B" & J, "G" & J + 5).Font.Bold = True
        Range("C" & J, "G" & J + 5).NumberFormat = "0.00"
        J = J + 7
    Loop
    
'Format columns
    Columns("H:I").Delete
    Columns("A:G").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Selection.Replace What:="0", Replacement:="", LookAt:=xlWhole
    Selection.EntireColumn.AutoFit
    Columns("H:H").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Range("A1").Select
    
'Insert Header
    Range("I1") = "=Partic!J2"
    Range("I1").NumberFormat = "m/d/yyyy"
    ActiveSheet.PageSetup.CenterHeader = "Contributions and Loans " & Range("I1")
    Range("I1") = ""
End If
End Sub
