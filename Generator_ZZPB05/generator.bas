Sub generator_zzpb05()

'lock screenupadting and change calculation to manual calculation
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'declare ALL variables
Dim wsname As String 'with "pro-zre"
Dim wsname2 As String 'output sheet
Dim wsname4 As String 'with "zgl-pro"
Dim wsname_source As String
Dim rw As Long, rw2 As Long, rw3 As Integer, rw4 As Long
Dim kom As Range, kom2 As Range, kom3 As Range
Dim tmpi As Integer, tmpj As Integer, tmpk As Integer
Dim sek As Integer, los_sek1 As Integer, los_sek2 As Integer, los_sek3 As Integer
Dim dt As Date

'set name of using sheets, base on current date & time
wsname2 = Year(Date) & Format(Month(Date), "00") & Format(Day(Date), "00") & "_" _
  & Format(Hour(Time), "00") & Format(Minute(Time), "00") & Format(Second(Time), "00")
wsname = wsname2 & "_pro-zre"
wsname4 = wsname2 & "_zgl-pro"
wsname_source = ActiveSheet.Name

'rw is number of rows in the source sheet
rw = Range("A1").End(xlDown).Row

'filtr right data basing on assumptions and copy them
'unable to filtr 3 criterias where all of them have "<>" inside
With Worksheets(wsname_source).Range("B7:F" & rw)
  .AutoFilter Field:=5, Criteria1:="<>PaK Zdrowie", Operator:=xlAnd, Criteria2:="<>portal świadczeniodawcy"
  .AutoFilter Field:=2, Criteria1:="ZGŁOSZENIE"
  .AutoFilter Field:=3, Criteria1:="PROPOZYCJA"
  .Copy
End With

'create new sheet and paste only necessary data; delete unnecessary columns and change date format
Sheets.Add After:=ActiveSheet
ActiveSheet.Name = wsname4
With Worksheets(wsname4)
  .Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
  .Columns("B:C").Delete
  .Range("B:B").NumberFormat = "m/d/yyyy"
  .Range("B:B").Replace What:=" *", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows
End With

'rw4 is number of rows in new sheet
rw4 = Worksheets(wsname4).Range("A1").End(xlDown).Row

'filtr the last criteria and delete output rows (with headers), then insert new headers with 1 new column
With Worksheets(wsname4)
  .Range("A1:C" & rw4).AutoFilter Field:=3, Criteria1:="ass-system"
  .Range(Range("A1"), Range("A1").End(xlToRight).End(xlDown)).EntireRow.Delete
  .Range("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
  .Range("A1") = "CASE_NUMBER"
  .Range("B1") = "DATE"
  .Range("C1") = "LOGIN"
  .Range("A1:C" & rw4).RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes
End With

'rw4 update
rw4 = Worksheets(wsname4).Range("A1").End(xlDown).Row
'--------------------------------------------------------------------


'filtr right data basing on assumptions and copy them
'unable to filtr 3 criterias where all of them have "<>" inside
With Worksheets(wsname_source).Range("B7:F" & rw)
  .AutoFilter Field:=5, Criteria1:="<>PaK Zdrowie", Operator:=xlAnd, Criteria2:="<>portal świadczeniodawcy"
  .AutoFilter Field:=2, Criteria1:="PROPOZYCJA"
  .AutoFilter Field:=3, Criteria1:="ZREALIZOWANA"
  .Copy
End With

'create new sheet and paste only necessary data; delete unnecessary columns and change date format
Sheets.Add After:=ActiveSheet
ActiveSheet.Name = wsname
With Worksheets(wsname)
  .Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
  .Columns("B:C").Delete
  .Range("B:B").NumberFormat = "m/d/yyyy"
  .Range("B:B").Replace What:=" *", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows
End With

'rw2 is number of rows in new sheet
rw2 = Worksheets(wsname).Range("A1").End(xlDown).Row

'filtr the last criteria and delete output rows (with headers), then insert new headers with 1 new column
With Worksheets(wsname)
  .Range("A1:C" & rw2).AutoFilter Field:=3, Criteria1:="ass-system"
  .Range(Range("A1"), Range("A1").End(xlToRight).End(xlDown)).EntireRow.Delete
  .Range("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
  .Range("A1") = "CASE_NUMBER"
  .Range("B1") = "DATE"
  .Range("C1") = "LOGIN"
  .Range("D1") = "RANDOM_NUMBER"
  .Range("E1") = "IF_ZGL"
  .Range("A1:C" & rw2).RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes
End With

'rw2 update
rw2 = Worksheets(wsname).Range("A1").End(xlDown).Row

'fill new column with random numbers
For Each kom In Worksheets(wsname).Range("D2:D" & rw2)
  kom = Rnd
Next

'sort data based on the new colum (with random numbers)
With Worksheets(wsname).Sort
  .SortFields.Clear
  .SortFields.Add Key:=Range("D1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  .SetRange Range("A2:D" & rw2)
  .Header = xlNo
  .MatchCase = False
  .Orientation = xlTopToBottom
  .SortMethod = xlPinYin
  .Apply
End With

'add columns with uniqe values composed of case_number and login
For Each kom In Worksheets(wsname).Range("E2:E" & rw2)
  kom.Value = kom.Offset(0, -4) & kom.Offset(0, -2)
Next
For Each kom In Worksheets(wsname4).Range("E2:E" & rw4)
  kom.Value = kom.Offset(0, -4) & kom.Offset(0, -2)
Next

'look if for the case number exists the same login in zgl-pro
tmpi = 0
For Each kom In Worksheets(wsname).Range("E2:E" & rw2)
  tmpi = WorksheetFunction.CountIf(Worksheets(wsname4).Range("E2:E" & rw4), kom.Value)
  If tmpi > 0 Then
    kom.Value = "YES"
  Else: kom.Value = "NO"
  End If
Next

'copy only 2 columns from data and remove duplicates
With Worksheets(wsname)
  .Columns("B:C").Copy
  .Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
  .Range("F:G").RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
End With

'rw3 is number of rows in f:g columns
rw3 = Worksheets(wsname).Range("F1").End(xlDown).Row

'order by login, date
With Worksheets(wsname).Sort
  .SortFields.Clear
  .SortFields.Add Key:=Range("G2:G" & rw3), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  .SortFields.Add Key:=Range("F2:F" & rw3), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
  .SetRange Range("F1:G" & rw3)
  .Header = xlYes
  .MatchCase = False
  .Orientation = xlTopToBottom
  .SortMethod = xlPinYin
  .Apply
End With

'add new sheet - it is the output sheet
Sheets.Add After:=ActiveSheet
ActiveSheet.Name = wsname2

'copy logins to output sheet and remove duplicates
Worksheets(wsname).Range("G:G").Copy
Worksheets(wsname2).Range("A1").PasteSpecial Paste:=xlPasteValues
Worksheets(wsname2).Range("A1:A" & rw2).RemoveDuplicates Columns:=Array(1), Header:=xlYes

'name new columns
With Worksheets(wsname2)
  .Range("B1") = "CASE_NUMBER_1"
  .Range("C1") = "CASE_NUMBER_2"
  .Range("D1") = "CASE_NUMBER_3"
  .Range("E1") = "ZGL-PRO-ZRE"
  .Range("F1") = "CASES_TOTAL_NUMBER"
  .Range("G1") = "DIFFERENT_DATES"
End With

'init variables to loops
tmpi = 2
tmpj = 0
tmpk = 2

'loop to fill columns "CASES_TOTAL_NUMBER", "DIFFERENT_DATES" and all dates
Do While Worksheets(wsname2).Cells(tmpi, 1) <> ""
  With Worksheets(wsname2)
    .Cells(tmpi, 7) = WorksheetFunction.CountIf(Worksheets(wsname).Range("G:G"), Worksheets(wsname2).Cells(tmpi, 1))
    .Cells(tmpi, 6) = WorksheetFunction.CountIf(Worksheets(wsname).Range("C:C"), Worksheets(wsname2).Cells(tmpi, 1))
    .Cells(tmpi, 5) = WorksheetFunction.CountIfs(Worksheets(wsname).Range("C:C"), Worksheets(wsname2).Cells(tmpi, 1), Worksheets(wsname).Range("E:E"), "YES")
  End With
  tmpj = Worksheets(wsname2).Cells(tmpi, 7).Value
  Do While tmpj > 0
    Worksheets(wsname2).Cells(tmpi, 7 + tmpj) = Worksheets(wsname).Cells(tmpk, 6)
    tmpj = tmpj - 1
    tmpk = tmpk + 1
  Loop
  tmpi = tmpi + 1
Loop

rw3 = Worksheets(wsname2).Range("A1").End(xlDown).Row
For Each kom In Worksheets(wsname2).Range("B2:B" & rw3)
  'sek => different_dates
  sek = kom.Offset(0, 5).Value
  If sek >= 3 Then
    los_sek1 = WorksheetFunction.RandBetween(1, sek / 3)
    los_sek2 = WorksheetFunction.RandBetween((sek / 3) + 1, sek / 3 * 2)
    los_sek3 = WorksheetFunction.RandBetween((sek / 3 * 2) + 1, sek)
    
    'Case 1
    For Each kom2 In Worksheets(wsname).Range("C2:C" & rw2)
      If kom2 = kom.Offset(0, -1) And kom.Offset(0, 5 + los_sek1) = kom2.Offset(0, -1) Then
        kom = kom2.Offset(0, -2)
        If kom2.Offset(0, 2) <> "YES" Then
          kom.Interior.ColorIndex = 3
        End If
        Exit For
      End If
    Next
    
    'Case 2
    For Each kom2 In Worksheets(wsname).Range("C2:C" & rw2)
      If kom2 = kom.Offset(0, -1) And kom.Offset(0, 5 + los_sek2) = kom2.Offset(0, -1) Then
        kom.Offset(0, 1) = kom2.Offset(0, -2)
        If kom2.Offset(0, 2) <> "YES" Then
          kom.Offset(0, 1).Interior.ColorIndex = 3
        End If
        Exit For
      End If
    Next
    
    'Case 3
    For Each kom2 In Worksheets(wsname).Range("C2:C" & rw2)
      If kom2 = kom.Offset(0, -1) And kom.Offset(0, 5 + los_sek3) = kom2.Offset(0, -1) Then
        kom.Offset(0, 2) = kom2.Offset(0, -2)
        If kom2.Offset(0, 2) <> "YES" Then
          kom.Offset(0, 2).Interior.ColorIndex = 3
        End If
        Exit For
      End If
    Next
    
    'Case when all Cases <> YES
    If kom.Interior.ColorIndex = 3 And kom.Offset(0, 1).Interior.ColorIndex = 3 And kom.Offset(0, 2).Interior.ColorIndex = 3 And kom.Offset(0, 3) > 0 Then
      For Each kom2 In Worksheets(wsname).Range("C2:C" & rw2)
        If kom2 = kom.Offset(0, -1) And kom2.Offset(0, 2) = "YES" Then
          'dt => Date when Case = "YES"
          dt = kom2.Offset(0, -1)
          'look for the right sector and put this case number to right column
          For Each kom3 In Range(Cells(kom.Row, 8), Cells(kom.Row, 7 + kom.Offset(0, 5).Value))
            If kom3 = dt Then
              Select Case (kom3.Column - 7)
              Case Is <= sek / 3
                  kom = kom2.Offset(0, -2)
                  kom.Interior.ColorIndex = 2
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
    
    ElseIf sek = 2 Then
      'Case 1
      For Each kom2 In Worksheets(wsname).Range("C2:C" & rw2)
        If kom2 = kom.Offset(0, -1) And kom.Offset(0, 6) = kom2.Offset(0, -1) Then
          kom = kom2.Offset(0, -2)
          If kom2.Offset(0, 2) <> "YES" Then
            kom.Interior.ColorIndex = 3
          End If
          Exit For
        End If
      Next
      
      'Case 2
      For Each kom2 In Worksheets(wsname).Range("C2:C" & rw2)
        If kom2 = kom.Offset(0, -1) And kom.Offset(0, 7) = kom2.Offset(0, -1) Then
          kom.Offset(0, 1) = kom2.Offset(0, -2)
          If kom2.Offset(0, 2) <> "YES" Then
            kom.Offset(0, 1).Interior.ColorIndex = 3
          End If
          Exit For
        End If
      Next
      
      'Case when both Cases <> YES
      If kom.Interior.ColorIndex = 3 And kom.Offset(0, 1).Interior.ColorIndex = 3 And kom.Offset(0, 3) > 0 Then
        For Each kom2 In Worksheets(wsname).Range("C2:C" & rw2)
          If kom2 = kom.Offset(0, -1) And kom2.Offset(0, 2) = "YES" Then
            'dt => Date when Case = "YES"
            dt = kom2.Offset(0, -1)
            'look for the right date and put this case number to right column
            If dt = kom.Offset(0, 6) Then
              kom = kom2.Offset(0, -2)
              kom.Interior.ColorIndex = 2
            Else:
              kom.Offset(0, 1) = kom2.Offset(0, -2)
              kom.Offset(0, 1).Interior.ColorIndex = 2
            End If
            Exit For
            
          End If
        Next
      End If
    
    Else:
      'Case 1
      For Each kom2 In Worksheets(wsname).Range("C2:C" & rw2)
        If kom2 = kom.Offset(0, -1) And kom.Offset(0, 3) > 0 Then
          If kom2.Offset(0, 2) = "YES" Then
            kom = kom2.Offset(0, -2)
            Exit For
          End If
        ElseIf kom2 = kom.Offset(0, -1) Then
          kom = kom2.Offset(0, -2)
          kom.Interior.ColorIndex = 3
        End If
      Next
    
  End If
Next

'delete unnecessary sheets - leave only source and output
Application.DisplayAlerts = False
Worksheets(wsname).Delete
Worksheets(wsname4).Delete
Application.DisplayAlerts = True

'turn off autofilter in source sheet
If Worksheets(wsname_source).AutoFilterMode Then Worksheets(wsname_source).AutoFilterMode = False

'unlock screenupadting and change calculation to automatic calculation
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub