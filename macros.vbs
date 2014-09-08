Sub LastRow()
Dim lastWrittenRow, i As Integer

    For i = 1 To ActiveSheet.Cells("2", "J").Value
    lastWrittenRow = Cells(Rows.Count, "A").End(xlUp).Row
        ActiveSheet.Cells(lastWrittenRow + 1, 1).Value = "a"
        ActiveSheet.Cells(lastWrittenRow + 2, 1).Value = "b"
        ActiveSheet.Cells(lastWrittenRow + 3, 1).Value = "c"
    Next i

End Sub

--------------------------------------------------------------

Private Sub Worksheet_Change(ByVal Target As Excel.Range)
  If Target.Address = "$D$2" Then
   Dim lastWrittenRow1, lastWrittenRow2, i As Integer

    Range("A6:F99").Clear

    For i = 1 To ActiveSheet.Cells("2", "D").Value
        lastWrittenRow1 = Cells(Rows.Count, "C").End(xlUp).Row
        
        lastWrittenRow2 = Cells(Rows.Count, "B").End(xlUp).Row
            
            ActiveSheet.Cells(lastWrittenRow1 + 1, 1).Value = "R"
            ActiveSheet.Cells(lastWrittenRow1 + 1, 2).Value = ActiveSheet.Cells(lastWrittenRow2, 2).Value + 1
            ActiveSheet.Cells(lastWrittenRow1 + 1, 3).Value = "Ilość obwodów"
            ActiveSheet.Cells(lastWrittenRow1 + 1, 5).BorderAround _
                ColorIndex:=0, Weight:=xlThin
            ActiveSheet.Cells(lastWrittenRow1 + 2, 3).Value = "Ilość stref grzewczych"
            ActiveSheet.Cells(lastWrittenRow1 + 2, 5).BorderAround _
                ColorIndex:=0, Weight:=xlThin
            ActiveSheet.Cells(lastWrittenRow1 + 3, 3).Value = "Najodleglejszy termostat"
            ActiveSheet.Cells(lastWrittenRow1 + 3, 5).BorderAround _
                ColorIndex:=0, Weight:=xlThin
            ActiveSheet.Cells(lastWrittenRow1 + 3, 6).Value = "m"
    
    Next i
  End If
 End Sub
--------------------------------------------------------------

Sub podziel()
Dim kom As Range, zakres As Range

Set zakres = Selection

For Each kom In zakres
    kom.Value = kom.Value * 0.156134
Next

End Sub

------------------------------------------------

Sub wyczyść()
Dim kom As Range, zakres As Range

Set zakres = Selection

For Each kom In zakres
    If kom.Value <= 0 Then
        kom.Value = ""
    End If
Next

End Sub

------------------------------------------------

Sub PLN2EUR()
Dim kom As Range, zakres As Range


Set zakres = Selection

For Each kom In zakres
    kom.Value = kom.Value / 4
Next

End Sub

------------------------------------------------

Sub EUR2PLN()
Dim kom As Range, zakres As Range


Set zakres = Selection

For Each kom In zakres
    kom.Value = kom.Value * 4
Next

End Sub

------------------------------------------------

Sub kopia()
Dim kom As Range, zakres As Range
Dim i As Integer
Dim x As Integer


Set zakres = Selection
x = 1

For Each kom In zakres
    For i = 1 To 12
    Worksheets("3").Cells(x, 1).Value = kom.Value
    x = x + 1
    Next
    x = x + 1
     
    
Next

End Sub

--------------------------------------------------------------

Sub ColorValuesInSelectionByRange()
Dim kom As Range, zakres As Range

If Not TypeName(Selection) = "Range" Then
    Exit Sub
End If

Set zakres = Selection

For Each kom In zakres
    If kom.Value < 1000 Then
        kom.Font.Color = RGB(0, 0, 250)
        kom.Font.Bold = True
    End If
    If kom.Value > 1000 And kom.Value < 2000 Then
        kom.Font.Color = RGB(250, 0, 0)
        kom.Font.Bold = True
    End If
    If kom.Value > 2000 And kom.Value < 3000 Then
        kom.Font.Color = RGB(0, 250, 0)
        kom.Font.Bold = True
    End If
    If kom.Value > 3000 And kom.Value < 4000 Then
        kom.Font.Color = RGB(0, 125, 0)
        kom.Font.Bold = True
    End If
    If kom.Value > 4000 And kom.Value < 5000 Then
        kom.Font.Color = RGB(125, 0, 0)
        kom.Font.Bold = True
    End If
    If kom.Value > 5000 And kom.Value < 6000 Then
        kom.Font.Color = RGB(0, 0, 125)
        kom.Font.Bold = True
    End If
Next

End Sub

-----------------------------------------------------------

Sub kopiuj_wiersz_kolumna()

Application.ScreenUpdating = False

Dim x As Long
Dim y As Long
Dim i As Long

x = 1
y = 4
i = 1

For i = 1 To 1000

    Sheets("1").Select
    Cells(x, 1).Select
    Selection.Copy
    Sheets("2").Select
    Cells(i, 1).Select
    ActiveSheet.Paste
    
    Sheets("1").Select
    Cells(y, 1).Select
    Selection.Copy
    Sheets("2").Select
    Cells(i, 2).Select
    ActiveSheet.Paste

    x = x + 4
    y = y + 4
    
Next

End Sub

-------------------------------------------------------
'podmiana warosci ujemnych na 0
Sub put0()
Dim kom As Range, zakres As Range
Dim I As Integer

Set zakres = Selection
I = 0
For Each kom In zakres
    If kom.Value < 0 Then
        kom.Value = 0
        I = I + 1
    End If
    Next
Range("A1").Value = I
End Sub

--------------------------------------------------------
There are many times when it would be great to have any macro run at a predetermined time or run it at specified intervals. Fortunately Excel has made this a relatively simple task, when you know how.

Application.OnTime

This Method is what we can use to achieve the automatically running of Excel Macros.

Public dTime As Date
Dim lNum As Long

Sub RunOnTime()
    dTime = Now + TimeSerial(0, 0, 10)
    Application.OnTime dTime, "RunOnTime"
    
    lNum = lNum + 1
    If lNum = 3 Then
        Run "CancelOnTime"
    Else
        MsgBox lNum
    End If
    
End Sub

Sub CancelOnTime()
    Application.OnTime dTime, "RunOnTime", , False
End Sub

--------------------------------------------------------

Automatyczne odświeżanie tabeli przestawnej


Tabele przestawne nie mają (poza odświeżaniem przy otwarciu zeszytu) opcji automatycznego odswieżania wraz ze zmianą danych źródłowych.

Pewnym rozwiązaniem może być zastosowanie kodu VBA zdarzenie uaktywnienia arkusza.

Kod oczywiście umieszczamy w module arkusza w którym znajduje się tabela.

Private Sub Worksheet_Activate()

Dim mysheet As Worksheet
    'definiujemy w którym arkuszu sa dane źródłowe
    Set mysheet = Sheets("Dane")

    Application.ScreenUpdating = False

    'przypisujemy adres zakresu w jakim są dane źródłowe
    '(przy założeniu, że komórka A1 znajduje się w zakresie danych źródłowych)
    myrange = mysheet.Name & "!" & _
    mysheet.Range("A1").CurrentRegion.Address(ReferenceStyle:=xlR1C1)

    'przypisujemy zakres danych źródłowych tabeli przestawnej
    ActiveSheet.PivotTableWizard SourceType:=xlDatabase, SourceData:=myrange

    'odświeżamy tabelę przestawną
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh

    'ukrywamy paski narzędzi związanych z tabelą przestawną
    ActiveWorkbook.ShowPivotTableFieldList = False
    Application.CommandBars("PivotTable").Visible = False

    Application.ScreenUpdating = True

End Sub
--------------------------------------------------------
Sub przesunięcie()
'ActiveCell.Offset(y,x) - wybranie nowej aktywnej komórki względem obecnie aktywnej komórki y w pionie x w poziomie y - dodatnie wartości w dół ujemne w górę, x dodatnie wartości w prawo ujemne wartości w lewo.
ActiveCell.Offset(3, 0).Select
Selection.Copy
ActiveCell.Offset(-1, 1).Select
ActiveSheet.Paste
ActiveCell.Offset(0, -1).Select
End Sub
--------------------------------------------------------
'zapisac jako plik excel z rozszerzeniem .xlam i dolaczyc jako dodatek do excel'
Function Słownie(Liczba As Variant, Optional CzyWaluta) As Variant

Dim LiczbaP, Wynik, Slowo, SlowoP, Slowo2, i, Przyrostki
Dim Przyrostek, Przedrostek, Grosze, Jednostki, dziesiatki, setki, gr

If IsMissing(CzyWaluta) Then CzyWaluta = True

If Liczba < 0 Then
Liczba = -Liczba
Przedrostek = "minus "
End If


Grosze = ""
If InStr(1, Liczba, ",", 1) > 0 Then
 Grosze = Right(Liczba, Len(Liczba) - InStr(1, Liczba, ",", 1))
 If Len(Grosze) = 1 Then Grosze = Grosze & "0"
 If Len(Grosze) > 2 Then Grosze = Left(Grosze, 2)
 Liczba = Left(Liczba, InStr(1, Liczba, ",", 1) - 1)
End If
Jednostki = Array("", "jeden", "dwa", "trzy", "cztery", _
                  "pięć", "sześć", "siedem", "osiem", "dziewięć", _
                  "dziesięć", "jedenaście", "dwanaście", "trzynaście", _
                  "czternaście", "piętnaście", "szesnaście", "siedemnaście", _
                  "osiemnaście", "dziewiętnaście")
dziesiatki = Array("", "dziesięć", "dwadzieścia", "trzydzieści", "czterdzieści", _
                  "pięćdziesiąt", "sześćdziesiąt", "siedemdziesiąt", _
                  "osiemdziesiąt", "dziewięćdziesiąt")
setki = Array("", "sto", "dwieście", "trzysta", "czterysta", "pięćset", "sześćset", _
              "siedemset", "osiemset", "dziewięćset")
Slowo = ""
For gr = 1 To 2
If Len(Liczba) - (Len(Liczba) \ 3) * 3 = 2 Then Liczba = "0" & Liczba
If Len(Liczba) - (Len(Liczba) \ 3) * 3 = 1 Then Liczba = "00" & Liczba
For i = 1 To (Len(Liczba) + 2) \ 3
  SlowoP = ""
  If i > 1 Then
    LiczbaP = Mid(Liczba, Len(Liczba) - (i * 3) + 1, 3)
  Else
    LiczbaP = Liczba
  End If
  If Right(LiczbaP, 2) < 20 Then
    SlowoP = Jednostki(Right(LiczbaP, 2)) & " " & SlowoP
  Else
    Slowo2 = dziesiatki(Left(Right(LiczbaP, 2), 1))
    Slowo2 = Slowo2 & " " & Jednostki(Right(LiczbaP, 1))
    SlowoP = Slowo2 & " " & SlowoP
  End If
  If LiczbaP > 99 Then
   SlowoP = setki(Left(Right(LiczbaP, 3), 1)) & " " & SlowoP
  End If
  Select Case i
   Case 1:
            If CzyWaluta Then
              If (gr = 2) Then
               Przyrostki = Array("grosz", "grosze", "groszy")
              Else
               Przyrostki = Array("złoty ", "złote ", "złotych ")
              End If
            Else
              If (gr = 2) Then
               Przyrostki = Array("setna", "setne", "setnych")
              Else
               Przyrostki = Array("", "", "")
              End If
            End If
   Case 2:  Przyrostki = Array("tysiąc ", "tysiące ", "tysięcy ")
   Case 3:  Przyrostki = Array("milion ", "miliony ", "milionów ")
   Case 4:  Przyrostki = Array("miliard ", "miliardy ", "miliardów ")
   Case 5:  Przyrostki = Array("bilion ", "biliony ", "bilionów ")
  End Select
  If ((LiczbaP <> 0) And i > 1) Or (gr > 0) Then
   If LiczbaP <> 0 Then
     If LiczbaP = 1 Then
      Przyrostek = Przyrostki(0)
     Else
        If ((Right(LiczbaP, 1) > 1) And (Right(LiczbaP, 1) < 5)) Or _
           ((Right(LiczbaP, 2) > 21) And (Right(LiczbaP, 1) > 1) And _
            (Right(LiczbaP, 1) < 5)) Then Przyrostek = Przyrostki(1)
        If ((Right(LiczbaP, 2) > 4) And (Right(LiczbaP, 2) < 22)) Or _
           ((Right(LiczbaP, 2) > 21) And (Right(LiczbaP, 1) > 4) And _
            (Right(LiczbaP, 1) < 22)) Or (Right(LiczbaP, 1) = 0) Or _
            (Right(LiczbaP, 1) = 1) Then Przyrostek = Przyrostki(2)
     End If
     If gr = 1 Then
      Slowo = SlowoP & Przyrostek & Slowo
     Else
      Slowo = Slowo & SlowoP & Przyrostek
     End If
   End If
  End If
Next i
If Grosze = "" Then
 Exit For
Else
 If Liczba > 0 Then If gr = 1 Then Slowo = Slowo & "i "
 Liczba = Grosze
End If
Next gr
If Liczba = 0 Then Slowo = "zero" & Slowo
Słownie = IIf(IsEmpty(Przedrostek), Slowo, Przedrostek & Slowo)
End Function

---------------------------------------------------

'Changing text box added to excel sheet (not user form) value and color (must use text box from developer controls in excel)

Private Sub Worksheet_Calculate()
Static OldVal As Variant
Dim i As Integer
  If Range("E77").Value <> OldVal Then
    For i = 1 To 19
        ActiveSheet.OLEObjects("TextBox" & i).Object.Text = Round(ActiveSheet.Cells(i + 76, 5) * 100, 0)
        If ActiveSheet.OLEObjects("TextBox" & i).Object.Text < 0 Then ActiveSheet.OLEObjects("TextBox" & i).Object.ForeColor = RGB(255, 0, 0)
        If ActiveSheet.OLEObjects("TextBox" & i).Object.Text >= 0 Then ActiveSheet.OLEObjects("TextBox" & i).Object.ForeColor = RGB(0, 176, 80)
        ActiveSheet.OLEObjects("TextBox" & i).Object.Text = ActiveSheet.OLEObjects("TextBox" & i).Object.Text & "%"
    Next i
  End If

End Sub

----------------------------------------------------------------

'Adding lines depended on specified value

Private Sub Worksheet_Change(ByVal Target As Excel.Range)
  If Target.Address = "$D$2" Then
   Dim lastWrittenRow1, lastWrittenRow2, i As Integer

    Range("A6:K99").Clear

    For i = 1 To ActiveSheet.Cells("2", "D").Value
        lastWrittenRow1 = Cells(Rows.Count, "C").End(xlUp).Row
        
        lastWrittenRow2 = Cells(Rows.Count, "B").End(xlUp).Row
            
            ActiveSheet.Cells(lastWrittenRow1 + 1, 1).Value = "R"
            ActiveSheet.Cells(lastWrittenRow1 + 1, 2).Value = ActiveSheet.Cells(lastWrittenRow2, 2).Value + 1
            ActiveSheet.Cells(lastWrittenRow1 + 1, 3).Value = "Ilość obwodów"
            ActiveSheet.Cells(lastWrittenRow1 + 1, 5).BorderAround _
                ColorIndex:=0, Weight:=xlThin
            ActiveSheet.Cells(lastWrittenRow1 + 2, 3).Value = "Ilość stref grzewczych"
            ActiveSheet.Cells(lastWrittenRow1 + 2, 5).BorderAround _
                ColorIndex:=0, Weight:=xlThin
            ActiveSheet.Cells(lastWrittenRow1 + 3, 3).Value = "Najodleglejszy termostat"
            ActiveSheet.Cells(lastWrittenRow1 + 3, 5).BorderAround _
                ColorIndex:=0, Weight:=xlThin
            ActiveSheet.Cells(lastWrittenRow1 + 3, 6).Value = "m"
    
    Next i
    
    ActiveSheet.Cells(lastWrittenRow1 + 5, 1).Value = "L.p."
    ActiveSheet.Cells(lastWrittenRow1 + 5, 2).Value = "Kod"
    ActiveSheet.Cells(lastWrittenRow1 + 5, 3).Value = "Nazwa"
    ActiveSheet.Cells(lastWrittenRow1 + 5, 8).Value = "Ilość"
    ActiveSheet.Cells(lastWrittenRow1 + 5, 9).Value = "J.m."
    ActiveSheet.Cells(lastWrittenRow1 + 5, 10).Value = "Cea netto"
    ActiveSheet.Cells(lastWrittenRow1 + 5, 11).Value = "Wartość netto"
    
    For i = 1 To 20
        Range(ActiveSheet.Cells(lastWrittenRow1 + (i + 4), 3), ActiveSheet.Cells(lastWrittenRow1 + (i + 4), 7)).Merge
    Next i
    
  End If
 End Sub
 
--------------------------------------------------------

Private Sub Worksheet_Calculate()
'przypisywanie wartosci w polach tekstowych excel (formatki activeX) oraz zmiana ich koloru w aktywnym arkuszu

Static OldVal1 As Variant
Static OldVal2 As Variant

Dim i As Integer
Dim j As Integer

  If Range("F27").Value <> OldVal1 Then
    For i = 1 To 19
        ActiveSheet.OLEObjects("TextBox" & i).Object.Text = Round(ActiveSheet.Cells(i + 26, 6) * 100, 0)
        If ActiveSheet.OLEObjects("TextBox" & i).Object.Text < 0 Then ActiveSheet.OLEObjects("TextBox" & i).Object.ForeColor = RGB(255, 0, 0)
        If ActiveSheet.OLEObjects("TextBox" & i).Object.Text >= 0 Then ActiveSheet.OLEObjects("TextBox" & i).Object.ForeColor = RGB(0, 176, 80)
        ActiveSheet.OLEObjects("TextBox" & i).Object.Text = ActiveSheet.OLEObjects("TextBox" & i).Object.Text & "%"
    Next i
  End If
  
  If Range("D99").Value <> OldVal2 Then
    For j = 20 To 32
        ActiveSheet.OLEObjects("TextBox" & j).Object.Text = Round(ActiveSheet.Cells(99, j - 16) * 100, 0)
        If ActiveSheet.OLEObjects("TextBox" & j).Object.Text < 0 Then ActiveSheet.OLEObjects("TextBox" & j).Object.ForeColor = RGB(255, 0, 0)
        If ActiveSheet.OLEObjects("TextBox" & j).Object.Text >= 0 Then ActiveSheet.OLEObjects("TextBox" & j).Object.ForeColor = RGB(0, 176, 80)
        ActiveSheet.OLEObjects("TextBox" & j).Object.Text = ActiveSheet.OLEObjects("TextBox" & j).Object.Text & "%"
    Next j
  End If
  
End Sub
--------------------------------------------------------

Private Sub CheckBox1_Change()
'wlaczanie lub wylaczanie kontrolki combobox na podstawie checkboxa 

  If CheckBox1.Value = True Then
    ActiveSheet.OLEObjects("ComboBox1").Object.Enabled = True
  Else: ActiveSheet.OLEObjects("ComboBox1").Object.Enabled = False
  End If
  
End Sub
--------------------------------------------------------

Private Sub Worksheet_Calculate()
'przypisywanie wartosci w polach tekstowych excel (formatki activeX) oraz zmiana ich koloru w wyznaczonym arkuszu (w przykładzie użyto arkusz o nazwie "Statystyka")
Static OldVal1 As Variant
Static OldVal2 As Variant

Dim i As Integer
Dim j As Integer

  If Range("F27").Value <> OldVal1 Then
    For i = 1 To 19
        Worksheets("Statystyka").OLEObjects("TextBox" & i).Object.Text = Round(Worksheets("Statystyka").Cells(i + 26, 6) * 100, 0)
        If Worksheets("Statystyka").OLEObjects("TextBox" & i).Object.Text < 0 Then Worksheets("Statystyka").OLEObjects("TextBox" & i).Object.ForeColor = RGB(255, 0, 0)
        If Worksheets("Statystyka").OLEObjects("TextBox" & i).Object.Text >= 0 Then Worksheets("Statystyka").OLEObjects("TextBox" & i).Object.ForeColor = RGB(0, 176, 80)
        Worksheets("Statystyka").OLEObjects("TextBox" & i).Object.Text = Worksheets("Statystyka").OLEObjects("TextBox" & i).Object.Text & "%"
    Next i
  End If

  If Range("D99").Value <> OldVal2 Then
    For j = 20 To 32
        Worksheets("Statystyka").OLEObjects("TextBox" & j).Object.Text = Round(Worksheets("Statystyka").Cells(99, j - 16) * 100, 0)
        If Worksheets("Statystyka").OLEObjects("TextBox" & j).Object.Text < 0 Then Worksheets("Statystyka").OLEObjects("TextBox" & j).Object.ForeColor = RGB(255, 0, 0)
        If Worksheets("Statystyka").OLEObjects("TextBox" & j).Object.Text >= 0 Then Worksheets("Statystyka").OLEObjects("TextBox" & j).Object.ForeColor = RGB(0, 176, 80)
        Worksheets("Statystyka").OLEObjects("TextBox" & j).Object.Text = Worksheets("Statystyka").OLEObjects("TextBox" & j).Object.Text & "%"
    Next j
  End If

End Sub
