Sub preparing_report()

Dim rw2 As Long

Dim rng As Range

rw2 = Range("A2").End(xlDown).Row

Set rng = Range("A1:F" & rw2)
rng.Interior.ColorIndex = xlNone
    
    'obramowanie
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    'nagłówek
    With Range("A1:F1").Font
        .Bold = True
        .ThemeColor = xlThemeColorDark1
    End With
    With Range("A1:F1").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 6697728
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

'dopasowanie i usunięcie kolumn
Columns("A:F").EntireColumn.AutoFit
Columns("G:Z").Delete
Columns("A:A").Delete

'zaznaczenie a1 dla estetyki i wyłączenie linii siatki
ActiveWindow.DisplayGridlines = False
Range("A1").Select


End Sub
