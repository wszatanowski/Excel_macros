Sub randomize_table()

'wyłączenie odświeżania ekranu i obliczania
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim rw2 As Long
Dim kom As Range
rw2 = Range("B2").End(xlDown).Row

Dim wsname_base As String

wsname_base = ActiveSheet.Name

For Each kom In Range("L2:L" & rw2)
    kom = Rnd
Next


    Worksheets(wsname_base).Sort.SortFields.Clear
    Worksheets(wsname_base).Sort.SortFields.Add Key:=Range("L1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(wsname_base).Sort
        .SetRange Range("A2:L" & rw2)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Worksheets(wsname_base).Columns("L:L").Delete
Worksheets(wsname_base).Range("A1").Select

'włączenie odświeżania ekranu i obliczania
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub
