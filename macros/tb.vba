Sub FormatTB()
'
' Formats TB and removes rows with $0.00
'
' Keyboard Shortcut: Ctrl+Shift+T
Dim I As Integer 'Total number of rows
                  
'Find how many rows the report has and delete $0.00 values
    I = 1
    Do While Range("A" & I) <> ""
        If Range("B" & I) = 0 Then
            Rows(I).Select
            Selection.Delete Shift:=xlUp
        Else
            I = I + 1
        End If
    Loop
    I = I - 1

'Format number column
    Columns("B:B").Select
    Selection.NumberFormat = "0.00"

'Delete unnecessary rows and columns
    Columns("D:AA").Select
    Selection.Delete Shift:=xlToLeft
    Rows(I + 1 & ":2000").Select
    Selection.Delete Shift:=xlUp

    Columns("A:C").EntireColumn.AutoFit
    Range("A1").Select


'Copy and paste values only
    'Range("A1:F" & I).Select
    'Selection.Copy
    'Selection.PasteSpecial Paste:=xlPasteValues
    
End Sub

Sub MarketFiles()
'
' MarketFiles Macro
'
' Keyboard Shortcut: Ctrl+Shift+M
'
Dim I As Integer 'Total number of rows

    Columns("A:J").Select
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .Bold = False
   End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
   
'Delete header rows
    Rows("1:5").Select
    Selection.Delete

'Find how many rows the report has and add missing Account+BP values
    I = 2
    Do While Range("A" & I) <> ""
        If Range("D" & I) = 0 Then
            Range("D" & I) = Range("A" & I)
        End If
        If Range("J" & I) = 0 Then
            Rows(I).Select
            Selection.Delete Shift:=xlUp
        Else
            I = I + 1
        End If
    Loop
    I = I - 1
    
    Columns("A:C").Select
    Selection.Delete
    Columns("G:G").Select
    Selection.NumberFormat = "0.00"
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("D:G").Select
    Selection.Delete
    Rows("1:1").Select
    Selection.Delete
    Columns("A:C").EntireColumn.AutoFit
    Range("A1").Select
    
End Sub

Sub MarketFiles()
'
' MarketFiles Macro
'
' Keyboard Shortcut: Ctrl+Shift+M
'
Dim I As Integer 'Total number of rows

    Columns("A:J").Select
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .Bold = False
   End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
   
'Delete header rows
    Rows("1:5").Select
    Selection.Delete

'Find how many rows the report has and add missing Account+BP values
    I = 2
    Do While Range("A" & I) <> ""
        If Range("D" & I) = 0 Then
            Range("D" & I) = Range("A" & I)
        End If
        If Range("J" & I) = 0 Then
            Rows(I).Select
            Selection.Delete Shift:=xlUp
        Else
            I = I + 1
        End If
    Loop
    I = I - 1
    
    Columns("A:C").Select
    Selection.Delete
    Columns("G:G").Select
    Selection.NumberFormat = "0.00"
    Selection.Cut
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight
    Columns("D:G").Select
    Selection.Delete
    Rows("1:1").Select
    Selection.Delete
    Columns("A:C").EntireColumn.AutoFit
    Range("A1").Select
    
End Sub
