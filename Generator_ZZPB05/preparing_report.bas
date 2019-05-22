Sub preparing_report()

Dim rw2 As Integer
Dim rng As Range

rw2 = Range("A2").End(xlDown).Row

Set rng = Range("A1:D" & rw2)
rng.Interior.ColorIndex = xlNone
    
    'make border
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
    
    'make header
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

'fit columns and delete unnecessary columns
Columns("A:D").EntireColumn.AutoFit
Columns("E:Z").Delete

'select A1 and hide gridlines
ActiveWindow.DisplayGridlines = False
Range("A1").Select

End Sub