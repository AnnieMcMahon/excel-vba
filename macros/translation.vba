Sub CreateTranslationTab()
'
' CreateTranslationTab Macro
'
' Keyboard Shortcut: Ctrl+Shift+W
'
Dim ReportName As String 'This workbook
Dim MacroName As String 'Macro workbook
Dim TabName As String 'Tab name
Dim Location As String 'First 3 digits of tab name
Dim I As Integer 'Total number of rows in TB
Dim J As Integer 'Counts BS and P&L rows
Dim K As Integer 'Multipurpose Counter
Dim L As Integer 'Row of each location

Dim CurrencyCode As String 'Currency to use
Dim PLacct As String 'Beginning of P&L accounts
Dim StockAcct As String 'Stock Acct #
Dim StockAmount 'Stock Amount
Dim REAcct As String 'Retained Earnings Acct #
Dim REAmount 'Retained Earnings Amount
Dim CTABegAmount 'CTA Beginning Amount

Dim RateCurrency 'Rate to use for BS items
Dim YearCurrency 'Rate to use for P&L items


'Copy original TB
    TabName = ActiveSheet.Name
    Sheets(TabName).Copy After:=Sheets(1)
    Sheets(Left(TabName, 27) & " (2)").Name = "Translation"
    TabName = "Translation"
    MacroName = "USD Translation Macro 2022.xlsm"
    ReportName = ActiveWorkbook.Name

'Find how many rows the report has
    I = 1
    Do While Range("A" & I) <> ""
        I = I + 1
    Loop
    I = I - 1
            
'Insert formulas
    Columns("B:B").Select
    Selection.NumberFormat = "0.00"
    Columns("D:K").Select
    Selection.Delete Shift:=xlToLeft
    Range("E1").FormulaR1C1 = "=ROUND((RC[-1]*RC[-3]),2)"
    Range("E1").Select
    Selection.AutoFill Destination:=Range("E1:E" & I + 2), Type:=xlFillDefault
    Columns("E:E").Select
    Selection.NumberFormat = "0.00"
            
'Determine currency rate, row, and P&L Accounts
Location = Left(ReportName, 3)
Select Case Location
'    Case "070", "071", "072", "073", "074", "075"
'        CurrencyCode = "CAD"
'        PLacct = "7"
    Case "301"
        CurrencyCode = "JPY"
        PLacct = "5"
    Case "311"
        CurrencyCode = "KRW"
        PLacct = "4"
    Case "321"
        CurrencyCode = "CNY"
        PLacct = "4"
    Case "351"
        CurrencyCode = "EUR"
        PLacct = "8"
    Case "352"
        CurrencyCode = "GBP"
        PLacct = "4"
    Case "353"
        CurrencyCode = "EUR"
        PLacct = "6"
End Select

'Store currency rate in variables
    Windows(MacroName).Activate
    K = 2
    Do While K < 8
        If Range("A" & K) = CurrencyCode Then
            RateCurrency = Range("D" & K)
            YearCurrency = Range("B" & K)
            K = 8
        Else
            K = K + 1
        End If
    Loop

'Store location row in variable L
    Windows(MacroName).Activate
    K = 2
    Do While K < 16
        If Range("F" & K) = Location Then
            L = K
            K = 16
        Else
            K = K + 1
        End If
    Loop

'Store information in variables
    StockAcct = Range("H" & L)
    StockAmount = Range("I" & L)
    REAcct = Range("J" & L)
    REAmount = Range("K" & L)
    
'Add currency rates for BS items
    Windows(ReportName).Activate
    J = 1
    Do While Range("A" & J) <> StockAcct And Range("A" & J) <> REAcct And Range("A" & J) <> ""
        Range("A" & J & ":C" & J).Select
        With Selection.Interior
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.799981688894314
        End With
        Range("D" & J) = RateCurrency
        J = J + 1
    Loop

'Add Stocks
    If Range("A" & J) = StockAcct Then
        Range("A" & J & ":E" & J).Select
        With Selection.Interior
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.599993896298105
        End With
        Range("E" & J) = StockAmount
        J = J + 1
    End If
    
'Add constant accounts for TW Spain
    If Location = "353" Then
        Range("A" & J & ":E" & J + 2).Select
        With Selection.Interior
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.599993896298105
        End With
        Range("E" & J) = 0
        J = J + 1
        Range("E" & J) = 0
        J = J + 1
        Range("E" & J) = 0
        J = J + 1
    End If

'Add Retained Earnings
    If Range("A" & J) = REAcct Then
        Range("A" & J & ":E" & J).Select
        With Selection.Interior
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.599993896298105
        End With
        Range("E" & J) = REAmount
        J = J + 1
    End If

'Add currency rates for BS items - TW Spain
    If Location = "353" Then
        Do While Range("A" & J) < 60000000 And Range("A" & J) <> ""
            Range("A" & J & ":C" & J).Select
            With Selection.Interior
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0.799981688894314
            End With
            Range("D" & J) = RateCurrency
            J = J + 1
        Loop
    End If

'Add currency rates for P&L items
    Do While Range("A" & J) <> ""
        Range("A" & J & ":C" & J).Select
        With Selection.Interior
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.799981688894314
        End With
        Range("D" & J) = YearCurrency
        J = J + 1
    Loop
    
'Add CTA Beginning Balance amount
    J = I + 1
    Columns("D:E").EntireColumn.AutoFit
    Range("A" & J).FormulaR1C1 = "341000"
    Range("C" & J).FormulaR1C1 = "CTA Beginning Balance"
    Windows(MacroName).Activate
    Range("L" & L).Copy
    Windows(ReportName).Activate
    Range("E" & J).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Range("A" & J & ":E" & J).Select
    With Selection.Interior
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599993896298105
    End With
    
'Add CTA - current year translation amount
    J = J + 1
    Range("A" & J) = "342000"
    Range("C" & J) = "CTA Current Year Translation"
    Range("E" & J) = "=SUM(R[-" & J - 1 & "]C:R[-1]C) * -1"
    
    Columns("A:E").EntireColumn.AutoFit
    Range("A" & J).Select
End Sub

Sub CalculateCanadaUSD()
'
' CalculateFixedItems Macro
'
' Change fixed items to USD - and currency rate for Canadian locations (temporary)
'
' Keyboard Shortcut: Ctrl+Shift+A
'
Dim ReportName As String 'This workbook
Dim MacroName As String 'Macro workbook
Dim TabName As String 'Tab name
Dim Location As String 'First 3 digits of tab name
Dim I As Integer 'Total number of rows in TB
Dim J As Integer 'Counts BS and P&L rows
Dim L As Integer 'Row of each location
Dim C As Integer 'Currency row
    
'Copy original TB
    TabName = ActiveSheet.Name
    Sheets(TabName).Copy After:=Sheets(1)
    Sheets(Left(TabName, 27) & " (2)").Name = "Translation"
    TabName = "Translation"
    MacroName = "USD Translation Macro 2022.xlsm"
    ReportName = ActiveWorkbook.Name

'Find how many rows the report has
    I = 1
    Do While Range("A" & I) <> ""
        I = I + 1
    Loop
    I = I - 1
            
'Insert formulas
    Columns("B:B").Select
    Selection.NumberFormat = "0.00"
    Columns("D:K").Select
    Selection.Delete Shift:=xlToLeft
    Range("E1").FormulaR1C1 = "=ROUND((RC[-1]*RC[-3]),2)"
    Range("E1").Select
    Selection.AutoFill Destination:=Range("E1:E" & I + 2), Type:=xlFillDefault
    Columns("E:E").Select
    Selection.NumberFormat = "0.00"

'Determine location
Location = Left(ReportName, 4)
Select Case Location
    Case "070M"
        L = 2
        C = 2
    Case "071T"
        L = 3
        C = 2
    Case "072V"
        L = 4
        C = 2
    Case "073H"
        L = 5
        C = 2
    Case "074C"
        L = 6
        C = 2
    Case "075C"
        L = 7
        C = 2
    Case "076Q"
        L = 8
        C = 2
    Case "070_"
        L = 9
        C = 2
End Select

'Add currency rates for BS items
    J = 1
    Do While Range("A" & J) < 60000000 And Range("A" & J) <> ""
        J = J + 1
    Loop
    J = J - 1
    Range("A1:C" & J).Select
    With Selection.Interior
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
    End With
    Windows(MacroName).Activate
    Range("D" & C).Select
    Selection.Copy
    Windows(ReportName).Activate
    Range("D1").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Selection.AutoFill Destination:=Range("D1:D" & J), Type:=xlFillDefault
    
'Add Stocks
    J = J + 1
    Do While Range("A" & J) < 62000000 And Range("A" & J) <> ""
        Range("A" & J & ":E" & J).Select
        With Selection.Interior
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.599993896298105
        End With
        If Range("A" & J) = 61000000 Then
            Windows(MacroName).Activate
            Range("I" & L).Select
            Selection.Copy
            Windows(ReportName).Activate
            Range("E" & J).Select
            Selection.PasteSpecial Paste:=xlPasteValues
        Else
            Range("E" & J) = 0
        End If
        J = J + 1
    Loop
    
'Add Retained Earnings
    If Range("A" & J) = 65000000 Then
        Range("A" & J & ":E" & J).Select
        With Selection.Interior
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.599993896298105
        End With
        Windows(MacroName).Activate
        Range("K" & L).Select
        Selection.Copy
        Windows(ReportName).Activate
        Range("E" & J).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        J = J + 1
    End If
    
'Add currency rates for P&L items
    Range("A" & J & ":C" & I).Select
    Application.CutCopyMode = False
    With Selection.Interior
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.799981688894314
    End With
    Windows(MacroName).Activate
    Range("B" & C).Select
    Selection.Copy
    Windows(ReportName).Activate
    Range("D" & J).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Selection.AutoFill Destination:=Range("D" & J & ":D" & I), Type:=xlFillDefault
    
'Add CTA Beginning Balance amount
    J = I + 1
    Columns("D:E").EntireColumn.AutoFit
    Range("A" & J).FormulaR1C1 = "341000"
    Range("C" & J).FormulaR1C1 = "CTA Beginning Balance"
    Windows(MacroName).Activate
    Range("L" & L).Copy
    Windows(ReportName).Activate
    Range("E" & J).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Range("A" & J & ":E" & J).Select
    With Selection.Interior
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599993896298105
    End With
    
'Add CTA - current year translation amount
    J = J + 1
    Range("A" & J) = "342000"
    Range("C" & J) = "CTA Current Year Translation"
    Range("E" & J) = "=SUM(R[-" & J - 1 & "]C:R[-1]C) * -1"
    
    Columns("A:E").EntireColumn.AutoFit
    Range("A" & J).Select
End Sub

Sub FinalTB()
'
' FinalTB Macro
'
' Keyboard Shortcut: Ctrl+Shift+F
'
    Sheets("Translation").Select
    Sheets("Translation").Copy After:=Sheets(2)
    Sheets("Translation (2)").Name = "Final TB USD"
    Columns("A:E").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Columns("E:E").Copy
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Columns("D:E").Delete
    Columns("A:C").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A1").Select
End Sub
