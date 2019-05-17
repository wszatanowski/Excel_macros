Sub generator()

'wyłączenie odświeżania ekranu i obliczania
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'deklarowanie zmiennych
Dim wsname As String 'z raportem
Dim wsname2 As String
Dim wsname_base As String 'oryginalny
Dim earliest As Date, latest As Date
Dim earstr As String, latstr As String
Dim rw1 As Byte, rw2 As Long, rw3 As Long

wsname_base = ActiveSheet.Name

'ustalenie minimalnej i maksymalnej daty
earliest = WorksheetFunction.Min(Range("B2:B100000"))
latest = WorksheetFunction.Max(Range("B2:B100000"))
earstr = Left(earliest, 10)
latstr = Left(latest, 10)

'oraz numerów wierszy: min, max i mid
rw1 = 2
rw2 = Range("B2").End(xlDown).Row
rw3 = (rw1 + rw2) / 2

'nazwa nowego arkusza
wsname2 = Year(Date) & Format(Month(Date), "00") & Format(Day(Date), "00") & "_" _
    & Format(Hour(Time), "00") & Format(Minute(Time), "00") & Format(Second(Time), "00")

wsname = wsname2 & "_raport"

'zrobienie nowego arkusza z raportami
Cells.Copy
Sheets.Add After:=ActiveSheet
ActiveSheet.Name = wsname
Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
Range("F:L").Delete
Range("A:A").Delete

Range("A:A").NumberFormat = "m/d/yyyy"
    Range("A:A").Replace What:=" *", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Worksheets(wsname).Range("A:D").RemoveDuplicates Columns:=Array(1, 2, 3, 4), _
    Header:=xlYes
Worksheets(wsname).Sort.SortFields.Clear
Worksheets(wsname).Sort.SortFields.Add Key _
    :=Range("D2:D1000000"), SortOn:=xlSortOnValues, Order:=xlAscending, _
    DataOption:=xlSortNormal
Worksheets(wsname).Sort.SortFields.Add Key _
    :=Range("C2:C1000000"), SortOn:=xlSortOnValues, Order:=xlAscending, _
    DataOption:=xlSortNormal
Worksheets(wsname).Sort.SortFields.Add Key _
    :=Range("A2:A1000000"), SortOn:=xlSortOnValues, Order:=xlDescending, _
    DataOption:=xlSortNormal
With Worksheets(wsname).Sort
    .SetRange Range("A1:D1000000")
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

Sheets.Add After:=ActiveSheet
ActiveSheet.Name = wsname2

'przeklejenie loginu, nazwiska i lidera
Worksheets(wsname).Range("B:D").Copy
Worksheets(wsname2).Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Application.CutCopyMode = False

Worksheets(wsname2).Range("A1:C" & rw2).RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes


Worksheets(wsname2).Sort.SortFields.Clear
    Worksheets(wsname2).Sort.SortFields.Add Key:=Range( _
        "C1:C2000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    Worksheets(wsname2).Sort.SortFields.Add Key:=Range( _
        "B1:B2000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With Worksheets(wsname2).Sort
        .SetRange Range("A1:C2000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'nazwa kolumn
Range("D1") = "NUMER_SPRAWY_1"
Range("E1") = "NUMER_SPRAWY_2"
Range("F1") = "NUMER_SPRAWY_3"
Range("G1") = "oiz"
Range("H1") = "wszystkie"
Range("I1") = "liczba_dni"

'wypełnianie liczby dni oraz dni w których dana osoba pracowała
Dim tmpi As Integer, tmpj As Integer
tmpi = 2
tmpj = 0
tmpk = 2
Do While Worksheets(wsname2).Cells(tmpi, 1) <> ""
    Worksheets(wsname2).Cells(tmpi, 9) = WorksheetFunction.CountIf(Worksheets(wsname).Range("B:B"), Worksheets(wsname2).Cells(tmpi, 1))
    Worksheets(wsname2).Cells(tmpi, 7) = WorksheetFunction.CountIfs(Worksheets(wsname_base).Range("C:C"), Worksheets(wsname2).Cells(tmpi, 1), Worksheets(wsname_base).Range("H:H"), "TAK")
    Worksheets(wsname2).Cells(tmpi, 8) = WorksheetFunction.CountIf(Worksheets(wsname_base).Range("C:C"), Worksheets(wsname2).Cells(tmpi, 1))
    tmpj = Worksheets(wsname2).Cells(tmpi, 9).Value
    Do While tmpj > 0
        Worksheets(wsname2).Cells(tmpi, 9 + tmpj) = Worksheets(wsname).Cells(tmpk, 1)
        tmpj = tmpj - 1
        tmpk = tmpk + 1
    Loop
    tmpi = tmpi + 1
Loop

'wypisywanie 1 sprawy
Dim kom As Range, kom2 As Range, kom3 As Range
For Each kom In Range("D2:D500")
    If kom.Offset(0, 5) = 1 And kom.Offset(0, 3) = 0 Then
        
        'dla: 1 pracujący, 0 zamkniętych
        For Each kom2 In Worksheets(wsname_base).Range("C2:C" & rw2)
            If kom2 = kom.Offset(0, -3) Then
                kom = kom2.Offset(0, -2)
                kom.Interior.ColorIndex = 3
                Exit For
            End If
        Next
        
        'dla: 1 pracujący, >0 zamkniętych
        ElseIf kom.Offset(0, 5) = 1 Then
            For Each kom2 In Worksheets(wsname_base).Range("C2:C" & rw2)
                If kom2 = kom.Offset(0, -3) And kom2.Offset(0, 5) = "TAK" Then
                    kom = kom2.Offset(0, -2)
                    Exit For
                End If
            Next
        
    End If
    If (kom.Offset(0, 5) = 2 Or kom.Offset(0, 5) = 3) And kom.Offset(0, 3) = 0 Then
        
        'dla: 2 albo 3 pracujący, 0 zamkniętych
        For Each kom2 In Worksheets(wsname_base).Range("C2:C" & rw2)
            If kom2 = kom.Offset(0, -3) And _
                kom.Offset(0, 6) = DateSerial(Year(kom2.Offset(0, -1)), Month(kom2.Offset(0, -1)), Day(kom2.Offset(0, -1))) Then
                    kom = kom2.Offset(0, -2)
                    kom.Interior.ColorIndex = 3
                    Exit For
            End If
        Next
        
        'dla: 2 albo 3 pracujący, >0 zamkniętych
    ElseIf kom.Offset(0, 5) = 2 Or kom.Offset(0, 5) = 3 Then
        For Each kom2 In Worksheets(wsname_base).Range("C2:C" & rw2)
            If kom2 = kom.Offset(0, -3) And kom2.Offset(0, 5) = "TAK" And _
                kom.Offset(0, 6) = DateSerial(Year(kom2.Offset(0, -1)), Month(kom2.Offset(0, -1)), Day(kom2.Offset(0, -1))) Then
                    kom = kom2.Offset(0, -2)
                Exit For
            End If
        Next
    End If
    
    'uzupełnienie jeśli >0 zamkniętych, ale nie danego dnia
    If kom.Offset(0, 5) <= 3 And kom.Offset(0, 3) <> 0 And kom = "" Then
        For Each kom2 In Worksheets(wsname_base).Range("C2:C" & rw2)
            If kom2 = kom.Offset(0, -3) And _
                kom.Offset(0, 6) = DateSerial(Year(kom2.Offset(0, -1)), Month(kom2.Offset(0, -1)), Day(kom2.Offset(0, -1))) Then
                    kom = kom2.Offset(0, -2)
                    kom.Interior.ColorIndex = 3
                    Exit For
            End If
        Next
    End If
    
Next

'wypisywanie 2 sprawy
For Each kom In Range("E2:E500")

    If (kom.Offset(0, 4) = 2 Or kom.Offset(0, 4) = 3) And kom.Offset(0, 2) = 0 Then
        
        'dla: 2 albo 3 pracujący, 0 zamkniętych
        For Each kom2 In Worksheets(wsname_base).Range("C2:C" & rw2)
            If kom2 = kom.Offset(0, -4) And _
                kom.Offset(0, 6) = DateSerial(Year(kom2.Offset(0, -1)), Month(kom2.Offset(0, -1)), Day(kom2.Offset(0, -1))) Then
                    kom = kom2.Offset(0, -2)
                    kom.Interior.ColorIndex = 3
                    Exit For
            End If
        Next
        
    'dla: 2 pracujący, >0 zamkniętych
    ElseIf kom.Offset(0, 4) = 2 Or kom.Offset(0, 4) = 3 Then
        If kom.Offset(0, -1).Interior.ColorIndex = 3 Then
            For Each kom2 In Worksheets(wsname_base).Range("C2:C" & rw2)
                If kom2 = kom.Offset(0, -4) And kom2.Offset(0, 5) = "TAK" And _
                    kom.Offset(0, 6) = DateSerial(Year(kom2.Offset(0, -1)), Month(kom2.Offset(0, -1)), Day(kom2.Offset(0, -1))) Then
                        kom = kom2.Offset(0, -2)
                        Exit For
                ElseIf kom2 = kom.Offset(0, -4) And kom.Offset(0, 4) = 3 And _
                    kom.Offset(0, 6) = DateSerial(Year(kom2.Offset(0, -1)), Month(kom2.Offset(0, -1)), Day(kom2.Offset(0, -1))) Then
                        kom = kom2.Offset(0, -2)
                        kom.Interior.ColorIndex = 3
                        Exit For
                End If
            Next
            
    'dla reszty uzupełnienie i zaznaczenie, jeśli nie zamknięta przez tę samą osobę
    Else:
        For Each kom2 In Worksheets(wsname_base).Range("C2:C" & rw2)
            If kom2 = kom.Offset(0, -4) And _
                kom.Offset(0, 6) = DateSerial(Year(kom2.Offset(0, -1)), Month(kom2.Offset(0, -1)), Day(kom2.Offset(0, -1))) Then
                    kom = kom2.Offset(0, -2)
                    If kom2.Offset(0, 5) <> "TAK" Then
                        kom.Interior.ColorIndex = 3
                    End If
                    Exit For
            End If
        Next
        
        
        End If
        
    End If
Next

'wypisywanie 3 sprawy
For Each kom In Range("F2:F500")

    If kom.Offset(0, 3) = 3 And kom.Offset(0, 1) = 0 Then
        
        'dla: 3 pracujący, 0 zamkniętych
        For Each kom2 In Worksheets(wsname_base).Range("C2:C" & rw2)
            If kom2 = kom.Offset(0, -5) And _
                kom.Offset(0, 6) = DateSerial(Year(kom2.Offset(0, -1)), Month(kom2.Offset(0, -1)), Day(kom2.Offset(0, -1))) Then
                    kom = kom2.Offset(0, -2)
                    kom.Interior.ColorIndex = 3
                    Exit For
            End If
        Next
        
    'dla: 3 pracujący, >0 zamkniętych, oba poprzednie <> TAK
    ElseIf kom.Offset(0, 3) = 3 Then
        If kom.Offset(0, -1).Interior.ColorIndex = 3 And kom.Offset(0, -2).Interior.ColorIndex = 3 Then
            For Each kom2 In Worksheets(wsname_base).Range("C2:C" & rw2)
                If kom2 = kom.Offset(0, -5) And kom2.Offset(0, 5) = "TAK" And _
                    kom.Offset(0, 6) = DateSerial(Year(kom2.Offset(0, -1)), Month(kom2.Offset(0, -1)), Day(kom2.Offset(0, -1))) Then
                        kom = kom2.Offset(0, -2)
                        Exit For
                End If
            Next
            
        'dla reszty uzupełnienie i zaznaczenie, jeśli nie zamknięta przez tę samą osobę
        Else:
            For Each kom2 In Worksheets(wsname_base).Range("C2:C" & rw2)
                If kom2 = kom.Offset(0, -5) And kom2.Offset(0, 5) = "TAK" And _
                    kom.Offset(0, 6) = DateSerial(Year(kom2.Offset(0, -1)), Month(kom2.Offset(0, -1)), Day(kom2.Offset(0, -1))) Then
                        kom = kom2.Offset(0, -2)
                        Exit For
                ElseIf kom2 = kom.Offset(0, -5) And _
                    kom.Offset(0, 6) = DateSerial(Year(kom2.Offset(0, -1)), Month(kom2.Offset(0, -1)), Day(kom2.Offset(0, -1))) Then
                        kom = kom2.Offset(0, -2)
                        kom.Interior.ColorIndex = 3
                        Exit For
                End If
            Next
        End If
        
        'jeśli poprzednie to błąd to wprowadzenie <> TAK, następnie przeszukujemy ponownie 1 i 2
        If kom = "" Then
            For Each kom2 In Worksheets(wsname_base).Range("C2:C" & rw2)
                If kom2 = kom.Offset(0, -5) And kom2.Offset(0, 5) <> "TAK" And _
                    kom.Offset(0, 6) = DateSerial(Year(kom2.Offset(0, -1)), Month(kom2.Offset(0, -1)), Day(kom2.Offset(0, -1))) Then
                        kom = kom2.Offset(0, -2)
                        kom.Interior.ColorIndex = 3
                        Exit For
                End If
            Next
            
            'sprawdzamy pierwszą sprawę
            For Each kom2 In Worksheets(wsname_base).Range("C2:C" & rw2)
                If kom2 = kom.Offset(0, -5) And kom2.Offset(0, 5) = "TAK" And _
                    kom.Offset(0, 6 - 2) = DateSerial(Year(kom2.Offset(0, -1)), Month(kom2.Offset(0, -1)), Day(kom2.Offset(0, -1))) Then
                        kom.Offset(0, -2) = kom2.Offset(0, -2)
                        kom.Offset(0, -2).Interior.ColorIndex = 2
                        Exit For
                End If
            Next
            
            'ewentualnie kończymy drugą sprawą
            If kom.Offset(0, -2).Interior.ColorIndex = 3 Then
                For Each kom2 In Worksheets(wsname_base).Range("C2:C" & rw2)
                    If kom2 = kom.Offset(0, -5) And kom2.Offset(0, 5) = "TAK" And _
                        kom.Offset(0, 6 - 1) = DateSerial(Year(kom2.Offset(0, -1)), Month(kom2.Offset(0, -1)), Day(kom2.Offset(0, -1))) Then
                            kom.Offset(0, -1) = kom2.Offset(0, -2)
                            kom.Offset(0, -1).Interior.ColorIndex = 2
                            Exit For
                    End If
                Next
            End If
            
        End If
        
    End If
Next


'clou tego makra - dla liczba dni > 3
'deklarujemy zmienne
Dim sek As Integer, los_sek1 As Integer, los_sek2 As Integer, los_sek3 As Integer

Dim dt As Date

For Each kom In Range("D2:D500")
    'zmienna = liczba dni
    sek = kom.Offset(0, 5).Value
    If sek > 3 Then
    
        los_sek1 = WorksheetFunction.RandBetween(1, sek / 3)
        los_sek2 = WorksheetFunction.RandBetween((sek / 3) + 1, sek / 3 * 2)
        los_sek3 = WorksheetFunction.RandBetween((sek / 3 * 2) + 1, sek)
        
        For Each kom2 In Worksheets(wsname_base).Range("C2:C" & rw2)
            'pierwsza sprawa
            If kom2 = kom.Offset(0, -3) And _
                kom.Offset(0, 5 + los_sek1) = DateSerial(Year(kom2.Offset(0, -1)), Month(kom2.Offset(0, -1)), Day(kom2.Offset(0, -1))) Then
                    kom = kom2.Offset(0, -2)
                    If kom2.Offset(0, 5) <> "TAK" Then
                        kom.Interior.ColorIndex = 3
                    End If
                    Exit For
            End If
        Next
            
        For Each kom2 In Worksheets(wsname_base).Range("C2:C" & rw2)
            'druga sprawa
            If kom2 = kom.Offset(0, -3) And _
                kom.Offset(0, 5 + los_sek2) = DateSerial(Year(kom2.Offset(0, -1)), Month(kom2.Offset(0, -1)), Day(kom2.Offset(0, -1))) Then
                    kom.Offset(0, 1) = kom2.Offset(0, -2)
                    If kom2.Offset(0, 5) <> "TAK" Then
                        kom.Offset(0, 1).Interior.ColorIndex = 3
                    End If
                    Exit For
            End If
        Next
       
        'jeśli dwa poprzednie negatywnie to szukamy w trzecim
        If kom.Interior.ColorIndex = 3 And kom.Offset(0, 1).Interior.ColorIndex = 3 Then
            For Each kom2 In Worksheets(wsname_base).Range("C2:C" & rw2)
                 'trzecia sprawa
                 If kom2 = kom.Offset(0, -3) And kom2.Offset(0, 5) = "TAK" And _
                     kom.Offset(0, 5 + los_sek3) = DateSerial(Year(kom2.Offset(0, -1)), Month(kom2.Offset(0, -1)), Day(kom2.Offset(0, -1))) Then
                         kom.Offset(0, 2) = kom2.Offset(0, -2)
                         Exit For
                 End If
             Next
        
        Else:
            For Each kom2 In Worksheets(wsname_base).Range("C2:C" & rw2)
                'trzecia sprawa jeśli nie było wcześniej dwóch negatywnych w poprzedniej
                If kom2 = kom.Offset(0, -3) And _
                    kom.Offset(0, 5 + los_sek3) = DateSerial(Year(kom2.Offset(0, -1)), Month(kom2.Offset(0, -1)), Day(kom2.Offset(0, -1))) Then
                        kom.Offset(0, 2) = kom2.Offset(0, -2)
                        If kom2.Offset(0, 5) <> "TAK" Then
                            kom.Offset(0, 2).Interior.ColorIndex = 3
                        End If
                        Exit For
                End If
            Next
        End If
        
        'jeżeli dwie pierwsze <> TAK, a w trzeciej nie ma, to uzupełnienie również <> TAK
        If kom.Offset(0, 2) = "" Then
            For Each kom2 In Worksheets(wsname_base).Range("C2:C" & rw2)
                If kom2 = kom.Offset(0, -3) And _
                    kom.Offset(0, 5 + los_sek3) = DateSerial(Year(kom2.Offset(0, -1)), Month(kom2.Offset(0, -1)), Day(kom2.Offset(0, -1))) Then
                        kom.Offset(0, 2) = kom2.Offset(0, -2)
                        'na wszelki wypadek tylko sprawdzam czy na pewno <> TAK
                        If kom2.Offset(0, 5) <> "TAK" Then
                            kom.Offset(0, 2).Interior.ColorIndex = 3
                        End If
                        Exit For
                End If
            Next
        End If
        
        'eliminowanie przypadku, w którym: liczba dni > 3, wszystkie <> TAK
        If kom.Interior.ColorIndex = 3 And kom.Offset(0, 1).Interior.ColorIndex = 3 And kom.Offset(0, 2).Interior.ColorIndex = 3 And _
            kom.Offset(0, 3) > 0 Then
            For Each kom2 In Worksheets(wsname_base).Range("C2:C" & rw2)
                If kom2 = kom.Offset(0, -3) And kom2.Offset(0, 5) = "TAK" Then
                    'zapisuję datę sprawy, która jest = TAK
                    dt = DateSerial(Year(kom2.Offset(0, -1)), Month(kom2.Offset(0, -1)), Day(kom2.Offset(0, -1)))
                    'sprawdzam, w którym sektorze należy umieścić sprawę i tam ją umieszczam
                    For Each kom3 In Range(Cells(kom.Row, 10), Cells(kom.Row, 10 + kom.Offset(0, 5).Value))
                        If kom3 = dt Then
                            Select Case (kom3.Column - 9)
                            Case Is <= sek / 3
                                kom.Offset = kom2.Offset(0, -2)
                                kom.Offset.Interior.ColorIndex = 2
                            Case Is > (sek / 3 * 2)
                                kom.Offset(0, 2) = kom2.Offset(0, -2)
                                kom.Offset(0, 1).Interior.ColorIndex = 2
                            Case Else
                                kom.Offset(0, 1) = kom2.Offset(0, -2)
                                kom.Offset(0, 1).Interior.ColorIndex = 2
                            End Select
                        End If
                    
                    Next
                    
                    Exit For
                End If
            Next
        End If
        
        
    End If
    
Next

'usunięcie arkusza roboczego
Application.DisplayAlerts = False
Worksheets(wsname).Delete
Application.DisplayAlerts = True

'włączenie odświeżania ekranu i obliczania
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub