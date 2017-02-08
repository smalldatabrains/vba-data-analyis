Public page As Worksheet

'Graph display in each ImageBox
Private Sub CheckBox1_Click()
Dim Table As PivotTable
Set Table = page.PivotTables("PivotTable2")
If CheckBox1.Value = True Then
Table.PivotFields("Location").CurrentPage = CheckBox1.Caption
Dim LeGraph As Chart
   Set LeGraph = page.ChartObjects(1).Chart
   NomImage = ThisWorkbook.Path & Application.PathSeparator & "Graf.gif"
   LeGraph.Export Filename:=NomImage, FilterName:="gif"
   UserForm1.Image1.Picture = LoadPicture(NomImage)
ElseIf CheckBox1.Value = False Then
UserForm1.Image1.Picture = LoadPicture("")
End If
End Sub
Private Sub CheckBox3_Click()
Dim Table As PivotTable
Set Table = page.PivotTables("PivotTable2")
If CheckBox3.Value = True Then
Table.PivotFields("Location").CurrentPage = CheckBox3.Caption
Dim LeGraph As Chart
   Set LeGraph = page.ChartObjects(1).Chart
   NomImage = ThisWorkbook.Path & Application.PathSeparator & "Graf.gif"
   LeGraph.Export Filename:=NomImage, FilterName:="gif"
   UserForm1.Image2.Picture = LoadPicture(NomImage)

Else
UserForm1.Image2.Picture = LoadPicture("")
End If
End Sub
Private Sub CheckBox4_Click()
Dim Table As PivotTable
Set Table = page.PivotTables("PivotTable2")
If CheckBox4.Value = True Then
Table.PivotFields("Location").CurrentPage = CheckBox4.Caption
Dim LeGraph As Chart
   Set LeGraph = page.ChartObjects(1).Chart
   NomImage = ThisWorkbook.Path & Application.PathSeparator & "Graf.gif"
   LeGraph.Export Filename:=NomImage, FilterName:="gif"
   UserForm1.Image3.Picture = LoadPicture(NomImage)

Else
UserForm1.Image3.Picture = LoadPicture("")
End If

End Sub


Private Sub CommandButton5_Click()
Call CommandButton7_Click
indexes = Array(1.5, 1.59, 1.6, 1.67, 1.74)
Dim i As Double

Dim compteur(4) As Double
i = 2
'calcul des valeurs de production et de rejets
While Sheets("eyebiz").Cells(i, 1).Value <> ""
Select Case Sheets("eyebiz").Cells(i, 10).Value
Case Is = indexes(0)
compteur(0) = compteur(0) + Sheets("eyebiz").Cells(i, 11).Value
Case Is = indexes(1)
compteur(1) = compteur(1) + Sheets("eyebiz").Cells(i, 11).Value
Case Is = indexes(2)
compteur(2) = compteur(2) + Sheets("eyebiz").Cells(i, 11).Value
Case Is = indexes(3)
compteur(3) = compteur(3) + Sheets("eyebiz").Cells(i, 11).Value
Case Is = indexes(4)
compteur(4) = compteur(4) + Sheets("eyebiz").Cells(i, 11).Value

End Select
i = i + 1
Wend

'affichage des labels
For j = 0 To UBound(indexes)
Set obj = Me.Controls.Add("forms.Label.1")
        With obj
            .Name = "monLabel" & j
            .Object.Caption = indexes(j)
            .Left = 590
            .Top = 15 * (j + 1) + 290
            .Width = 50
            .Height = 10
        End With
'calcul de la production
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 2
            .Object.Caption = compteur(j)
            .Left = 655
            .Top = 15 * (j + 1) + 290
            .Width = Label29.Width
            .Height = 10
        End With
'affiche les rejets
Sheets("surf").PivotTables("PivotTable2").PivotFields("Location").CurrentPage = "EYEBIZ1"
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = Sheets("indexes").Cells(j + 6, 2).Value
            .Left = 710
            .Top = 15 * (j + 1) + 290
            .Width = Label29.Width
            .Height = 10
        End With
' calcul le taux
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = Round(Sheets("indexes").Cells(j + 6, 2).Value / compteur(j) * 100, 2) & "%"
            .Left = 765
            .Top = 15 * (j + 1) + 290
            .Width = Label29.Width
            .Height = 10
        End With

Next j
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = "---------------------------------------------------------------------"
            .Left = 590
            .Top = 377
            .Width = 200
            .Height = 10
        End With
i = 2
j = 0
Erase compteur
'calcul des valeurs de production et de rejets
While Sheets("eolt").Cells(i, 1).Value <> ""
Select Case Sheets("eolt").Cells(i, 10).Value
Case Is = indexes(0)
compteur(0) = compteur(0) + Sheets("eolt").Cells(i, 11).Value
Case Is = indexes(1)
compteur(1) = compteur(1) + Sheets("eolt").Cells(i, 11).Value
Case Is = indexes(2)
compteur(2) = compteur(2) + Sheets("eolt").Cells(i, 11).Value
Case Is = indexes(3)
compteur(3) = compteur(3) + Sheets("eolt").Cells(i, 11).Value

End Select
i = i + 1
Wend

'affichage des labels
For j = 0 To UBound(indexes) - 1
Set obj = Me.Controls.Add("forms.Label.1")
        With obj
            .Name = "monLabel" & j
            .Object.Caption = indexes(j)
            .Left = 590
            .Top = 15 * (j + 1) + 375
            .Width = 50
            .Height = 10
        End With
'calcul de la production
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 2
            .Object.Caption = compteur(j)
            .Left = 655
            .Top = 15 * (j + 1) + 375
            .Width = Label29.Width
            .Height = 10
        End With
'affiche les rejets
Sheets("indexes").PivotTables("PivotTable2").PivotFields("Location").CurrentPage = "EOLT1"
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = Sheets("indexes").Cells(j + 6, 2).Value
            .Left = 710
            .Top = 15 * (j + 1) + 375
            .Width = Label29.Width
            .Height = 10
        End With
' calcul le taux
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = Round(Sheets("indexes").Cells(j + 6, 2).Value / compteur(j) * 100, 2) & "%"
            .Left = 765
            .Top = 15 * (j + 1) + 375
            .Width = Label29.Width
            .Height = 10
        End With

Next j
End Sub


Private Sub CommandButton6_Click()
Call CommandButton7_Click
surfs = Array("DS", "TDS", "Tradi")
Dim i As Double
Dim compteur(2) As Double
i = 2
'calcul des valeurs de production et de rejets
While Sheets("eyebiz").Cells(i, 1).Value <> ""
Select Case Sheets("eyebiz").Cells(i, 6).Value
Case Is = surfs(0)
compteur(0) = compteur(0) + Sheets("eyebiz").Cells(i, 11).Value
Case Is = surfs(1)
compteur(1) = compteur(1) + Sheets("eyebiz").Cells(i, 11).Value
Case Is = surfs(2)
compteur(2) = compteur(2) + Sheets("eyebiz").Cells(i, 11).Value
End Select
i = i + 1
Wend

'affichage des labels
For j = 0 To UBound(surfs)
Set obj = Me.Controls.Add("forms.Label.1")
        With obj
            .Name = "monLabel" & j
            .Object.Caption = surfs(j)
            .Left = 590
            .Top = 15 * (j + 1) + 290
            .Width = 50
            .Height = 10
        End With
'calcul de la production
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 2
            .Object.Caption = compteur(j)
            .Left = 655
            .Top = 15 * (j + 1) + 290
            .Width = Label29.Width
            .Height = 10
        End With
'affiche les rejets
Sheets("surf").PivotTables("PivotTable2").PivotFields("Location").CurrentPage = "EYEBIZ1"
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = Sheets("surf").Cells(j + 6, 7).Value
            .Left = 710
            .Top = 15 * (j + 1) + 290
            .Width = Label29.Width
            .Height = 10
        End With
' calcul le taux
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = Round(Sheets("surf").Cells(j + 6, 7).Value / compteur(j) * 100, 2) & "%"
            .Left = 765
            .Top = 15 * (j + 1) + 290
            .Width = Label29.Width
            .Height = 10
        End With

Next j
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = "---------------------------------------------------------------------"
            .Left = 590
            .Top = 345
            .Width = 200
            .Height = 10
        End With

i = 2
j = 0
Erase compteur
'calcul des valeurs de production et de rejets
While Sheets("eolt").Cells(i, 1).Value <> ""
Select Case Sheets("eolt").Cells(i, 6).Value
Case Is = surfs(0)
compteur(0) = compteur(0) + Sheets("eolt").Cells(i, 11).Value
Case Is = surfs(1)
compteur(1) = compteur(1) + Sheets("eolt").Cells(i, 11).Value
Case Is = surfs(2)
compteur(2) = compteur(2) + Sheets("eolt").Cells(i, 11).Value
End Select
i = i + 1
Wend

'affichage des labels
For j = 0 To UBound(surfs)
Set obj = Me.Controls.Add("forms.Label.1")
        With obj
            .Name = "monLabel" & j
            .Object.Caption = surfs(j)
            .Left = 590
            .Top = 15 * (j + 1) + 355
            .Width = 50
            .Height = 10
        End With
'calcul de la production
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 2
            .Object.Caption = compteur(j)
            .Left = 655
            .Top = 15 * (j + 1) + 355
            .Width = Label29.Width
            .Height = 10
        End With
'affiche les rejets
Sheets("surf").PivotTables("PivotTable2").PivotFields("Location").CurrentPage = "EOLT1"
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = Sheets("surf").Cells(j + 6, 6).Value
            .Left = 710
            .Top = 15 * (j + 1) + 355
            .Width = Label29.Width
            .Height = 10
        End With
' calcul le taux
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = Round(Sheets("surf").Cells(j + 6, 6).Value / compteur(j) * 100, 2) & "%"
            .Left = 765
            .Top = 15 * (j + 1) + 355
            .Width = Label29.Width
            .Height = 10
        End With

Next j
End Sub

Private Sub CommandButton7_Click()
Set obj = Me.Controls.Add("forms.Label.1")
        With obj
            .Name = "RAZ"
            .Object.Caption = ""
            .Left = 590
            .Top = 15 * (j + 1) + 290
            .Width = 220
            .Height = 800
        End With
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Label18_Click()

Dim j As Double
Dim rate2 As Double
Dim totaleolt As Double
j = 2
totaleolt = 0
While Sheets("eolt").Cells(j, 1).Value <> ""
totaleolt = totaleolt + Sheets("eolt").Cells(j, 11).Value
j = j + 1
Wend
rate2 = Round(5360 / totaleolt * 100, 2)
Label18.Caption = "Produced : " & totaleolt & Chr(10) & "Rejected : " & "5360" & Chr(10) & "Defect rate : " & rate2 & "%"

End Sub


Private Sub Label19_Click()
Dim j As Integer
Dim rate As Double
Dim totaleyebiz As Double
j = 2
totaleyebiz = 0
While Sheets("Eyebiz").Cells(j, 1).Value <> ""
totaleyebiz = totaleyebiz + Sheets("Eyebiz").Cells(j, 11).Value
j = j + 1
Wend
rate = Round(11917 / totaleyebiz * 100, 2)
Label19.Caption = "Produced : " & totaleyebiz & Chr(10) & "Rejected : " & "11917" & Chr(10) & "Defect rate : " & rate & "%"

End Sub


'KPI calculation for stock input
Private Sub CommandButton4_Click()
Call CommandButton7_Click
stocks = Array("Rx", "Stock")
Dim i As Double
Dim compteur(1) As Double
i = 2
'calcul des valeurs de production et de rejets
While Sheets("eyebiz").Cells(i, 1).Value <> ""
Select Case Sheets("eyebiz").Cells(i, 19).Value
Case Is = stocks(0)
compteur(0) = compteur(0) + Sheets("eyebiz").Cells(i, 11).Value
Case Is = stocks(1)
compteur(1) = compteur(1) + Sheets("eyebiz").Cells(i, 11).Value
End Select
i = i + 1
Wend

'affichage des labels
For j = 0 To 1
Set obj = Me.Controls.Add("forms.Label.1")
        With obj
            .Name = "monLabel" & j
            .Object.Caption = stocks(j)
            .Left = 590
            .Top = 15 * (j + 1) + 290
            .Width = 50
            .Height = 10
        End With
'calcul de la production
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 2
            .Object.Caption = compteur(j)
            .Left = 655
            .Top = 15 * (j + 1) + 290
            .Width = Label29.Width
            .Height = 10
        End With
'affiche les rejets
Sheets("stock").PivotTables("PivotTable2").PivotFields("Location").CurrentPage = "EYEBIZ1"
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = Sheets("stock").Cells(j + 6, 7).Value
            .Left = 710
            .Top = 15 * (j + 1) + 290
            .Width = Label29.Width
            .Height = 10
        End With
' calcul le taux
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = Round(Sheets("stock").Cells(j + 6, 7).Value / compteur(j) * 100, 2) & "%"
            .Left = 765
            .Top = 15 * (j + 1) + 290
            .Width = Label29.Width
            .Height = 10
        End With

Next j
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = "---------------------------------------------------------------------"
            .Left = 590
            .Top = 335
            .Width = 200
            .Height = 10
        End With
i = 2
j = 0
Erase compteur
'calcul des valeurs de production et de rejets
While Sheets("eolt").Cells(i, 1).Value <> ""
Select Case Sheets("eolt").Cells(i, 19).Value
Case Is = stocks(0)
compteur(0) = compteur(0) + Sheets("eolt").Cells(i, 11).Value
Case Is = stocks(1)
compteur(1) = compteur(1) + Sheets("eolt").Cells(i, 11).Value
End Select
i = i + 1
Wend

'affichage des labels
For j = 0 To 1
Set obj = Me.Controls.Add("forms.Label.1")
        With obj
            .Name = "monLabel" & j
            .Object.Caption = stocks(j)
            .Left = 590
            .Top = 15 * (j + 1) + 330
            .Width = 50
            .Height = 10
        End With
'calcul de la production
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 2
            .Object.Caption = compteur(j)
            .Left = 655
            .Top = 15 * (j + 1) + 330
            .Width = Label29.Width
            .Height = 10
        End With
'affiche les rejets
Sheets("stock").PivotTables("PivotTable2").PivotFields("Location").CurrentPage = "EOLT1"
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = Sheets("stock").Cells(j + 6, 6).Value
            .Left = 710
            .Top = 15 * (j + 1) + 330
            .Width = Label29.Width
            .Height = 10
        End With
' calcul le taux
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = Round(Sheets("stock").Cells(j + 6, 6).Value / compteur(j) * 100, 2) & "%"
            .Left = 765
            .Top = 15 * (j + 1) + 330
            .Width = Label29.Width
            .Height = 10
        End With

Next j
End Sub

'KPI calculation for shiftinput
Private Sub CommandButton2_Click()
Call CommandButton7_Click
shifts = Array("Shift1", "Shift2")
Dim i As Double
Dim compteur(1) As Double
i = 2
'calcul des valeurs de production et de rejets
While Sheets("eyebiz").Cells(i, 1).Value <> ""
Select Case Sheets("eyebiz").Cells(i, 2).Value
Case Is = shifts(0)
compteur(0) = compteur(0) + Sheets("eyebiz").Cells(i, 11).Value
Case Is = shifts(1)
compteur(1) = compteur(1) + Sheets("eyebiz").Cells(i, 11).Value
End Select
i = i + 1
Wend

'affichage des labels
For j = 0 To UBound(shifts)
Set obj = Me.Controls.Add("forms.Label.1")
        With obj
            .Name = "monLabel" & j
            .Object.Caption = shifts(j)
            .Left = 590
            .Top = 15 * (j + 1) + 290
            .Width = 50
            .Height = 10
        End With
'calcul de la production
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 2
            .Object.Caption = compteur(j)
            .Left = 655
            .Top = 15 * (j + 1) + 290
            .Width = Label29.Width
            .Height = 10
        End With
'affiche les rejets
Sheets("shift").PivotTables("PivotTable2").PivotFields("Location").CurrentPage = "EYEBIZ1"
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = Sheets("shift").Cells(j + 6, 7).Value
            .Left = 710
            .Top = 15 * (j + 1) + 290
            .Width = Label29.Width
            .Height = 10
        End With
' calcul le taux
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = Round(Sheets("shift").Cells(j + 6, 7).Value / compteur(j) * 100, 2) & "%"
            .Left = 765
            .Top = 15 * (j + 1) + 290
            .Width = Label29.Width
            .Height = 10
        End With

Next j

Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = "---------------------------------------------------------------------"
            .Left = 590
            .Top = 335
            .Width = 200
            .Height = 10
        End With
i = 2
j = 0
Erase compteur
'calcul des valeurs de production et de rejets
While Sheets("eolt").Cells(i, 1).Value <> ""
Select Case Sheets("eolt").Cells(i, 2).Value
Case Is = shifts(0)
compteur(0) = compteur(0) + Sheets("eolt").Cells(i, 11).Value
Case Is = shifts(1)
compteur(1) = compteur(1) + Sheets("eolt").Cells(i, 11).Value
End Select
i = i + 1
Wend

'affichage des labels
For j = 0 To 1
Set obj = Me.Controls.Add("forms.Label.1")
        With obj
            .Name = "monLabel" & j
            .Object.Caption = shifts(j)
            .Left = 590
            .Top = 15 * (j + 1) + 330
            .Width = 50
            .Height = 10
        End With
'calcul de la production
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 2
            .Object.Caption = compteur(j)
            .Left = 655
            .Top = 15 * (j + 1) + 330
            .Width = Label29.Width
            .Height = 10
        End With
'affiche les rejets
Sheets("shift").PivotTables("PivotTable2").PivotFields("Location").CurrentPage = "EOLT1"
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = Sheets("shift").Cells(j + 6, 6).Value
            .Left = 710
            .Top = 15 * (j + 1) + 330
            .Width = Label29.Width
            .Height = 10
        End With
' calcul le taux
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = Round(Sheets("shift").Cells(j + 6, 6).Value / compteur(j) * 100, 2) & "%"
            .Left = 765
            .Top = 15 * (j + 1) + 330
            .Width = Label29.Width
            .Height = 10
        End With

Next j
End Sub
'KPI calculation for baseinput A TERMINER MAIS LA BASE N EST PAS DANS MES LIGNES DE PROD
Private Sub CommandButton3_Click()
Call CommandButton7_Click
'base = Array("0.25", "0.5", "0.6", "0.75", "1", "1.25", "1.5", "1.75", "2", "2.25", "2.5", "2.75", "3", "3.25", "3.5", "3.6", "3.62", "3.75", "4", "4.25", "4.5", "4.75", "5", "5.25", "5.5", "5.75", "6", "6.25", "6.5", "6.75", "7", "7.25", "7.5", "8", "8.25", "8.5", "8.75", "9.5", "10")
Order = Array("HC-EDG", "MC-EDG", "TIN-EDG", "TIN-HC-EDG", "TIN-MC-EDG", "UNC-EDG")
Dim i As Double
Dim compteur(5) As Double
i = 2
'calcul des valeurs de production et de rejets
While Sheets("eyebiz").Cells(i, 1).Value <> ""
Select Case Sheets("eyebiz").Cells(i, 17).Value
Case Is = Order(0)
compteur(0) = compteur(0) + Sheets("eyebiz").Cells(i, 11).Value
Case Is = Order(1)
compteur(1) = compteur(1) + Sheets("eyebiz").Cells(i, 11).Value
Case Is = Order(2)
compteur(2) = compteur(2) + Sheets("eyebiz").Cells(i, 11).Value
Case Is = Order(3)
compteur(3) = compteur(3) + Sheets("eyebiz").Cells(i, 11).Value
Case Is = Order(4)
compteur(4) = compteur(4) + Sheets("eyebiz").Cells(i, 11).Value
Case Is = Order(5)
compteur(5) = compteur(5) + Sheets("eyebiz").Cells(i, 11).Value
End Select
i = i + 1
Wend

'affichage des labels
For j = 0 To UBound(Order)
Set obj = Me.Controls.Add("forms.Label.1")
        With obj
            .Name = "monLabel" & j
            .Object.Caption = Order(j)
            .Left = 590
            .Top = 15 * (j + 1) + 290
            .Width = 50
            .Height = 10
        End With
'calcul de la production
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 2
            .Object.Caption = compteur(j)
            .Left = 655
            .Top = 15 * (j + 1) + 290
            .Width = Label29.Width
            .Height = 10
        End With
'affiche les rejets
Sheets("order").PivotTables("PivotTable2").PivotFields("Location").CurrentPage = "EYEBIZ1"
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = Sheets("order").Cells(j + 6, 7).Value
            .Left = 710
            .Top = 15 * (j + 1) + 290
            .Width = Label29.Width
            .Height = 10
        End With
' calcul le taux
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = Round(Sheets("order").Cells(j + 6, 7).Value / compteur(j) * 100, 2) & "%"
            .Left = 765
            .Top = 15 * (j + 1) + 290
            .Width = Label29.Width
            .Height = 10
        End With
        
Next j
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = "---------------------------------------------------------------------"
            .Left = 590
            .Top = 395
            .Width = 200
            .Height = 10
        End With

i = 2
j = 0
Erase compteur
'calcul des valeurs de production et de rejet
While Sheets("eolt").Cells(i, 1).Value <> ""
Select Case Sheets("eolt").Cells(i, 17).Value
Case Is = Order(0)
compteur(0) = compteur(0) + Sheets("eolt").Cells(i, 11).Value
Case Is = Order(1)
compteur(1) = compteur(1) + Sheets("eolt").Cells(i, 11).Value
Case Is = Order(2)
compteur(2) = compteur(2) + Sheets("eolt").Cells(i, 11).Value
Case Is = Order(3)
compteur(3) = compteur(3) + Sheets("eolt").Cells(i, 11).Value
Case Is = Order(4)
compteur(4) = compteur(4) + Sheets("eolt").Cells(i, 11).Value
Case Is = Order(5)
compteur(5) = compteur(5) + Sheets("eolt").Cells(i, 11).Value
End Select
i = i + 1
Wend

'affichage des labels
For j = 0 To UBound(Order)
Set obj = Me.Controls.Add("forms.Label.1")
        With obj
            .Name = "monLabel" & j
            .Object.Caption = Order(j)
            .Left = 590
            .Top = 15 * (j + 1) + 400
            .Width = 50
            .Height = 10
        End With
'calcul de la production
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 2
            .Object.Caption = compteur(j)
            .Left = 655
            .Top = 15 * (j + 1) + 400
            .Width = Label29.Width
            .Height = 10
        End With
'affiche les rejets
Sheets("order").PivotTables("PivotTable2").PivotFields("Location").CurrentPage = "EOLT1"
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = Sheets("order").Cells(j + 6, 6).Value
            .Left = 710
            .Top = 15 * (j + 1) + 400
            .Width = Label29.Width
            .Height = 10
        End With
' calcul le taux
Set obj = Me.Controls.Add("forms.Label.1")
     With obj
            .Name = "monLabel" & j * 3
            .Object.Caption = Round(Sheets("order").Cells(j + 6, 6).Value / compteur(j) * 100, 2) & "%"
            .Left = 765
            .Top = 15 * (j + 1) + 400
            .Width = Label29.Width
            .Height = 10
        End With

Next j
End Sub



'manage views and user inputs
Private Sub MultiPage1_Change()
UserForm1.Image1.Picture = LoadPicture("")
UserForm1.Image2.Picture = LoadPicture("")
UserForm1.Image3.Picture = LoadPicture("")


Select Case MultiPage1.Value
Case Is = 0
Set page = Worksheets("indexes")
Case Is = 1
Set page = Worksheets("base")
Case Is = 2
Set page = Worksheets("shift")
Case Is = 3
Set page = Worksheets("stock")
Case Is = 4
Set page = Worksheets("side")
Case Is = 5
Set page = Worksheets("order")
Case Is = 6
Set page = Worksheets("surf")
Case Is = 7
Set page = Worksheets("index")
End Select

CheckBox1.Value = True
CheckBox4.Value = True
CheckBox3.Value = True
Call CheckBox1_Click
Call CheckBox3_Click
Call CheckBox4_Click
Call Label19_Click
Call Label18_Click

End Sub

' Display the initial page
Private Sub UserForm_Initialize()
With Me
        .StartUpPosition = 3
        .Width = Application.Width - 10
        .Height = Application.Height
        .Left = 0
        .Top = 0
End With

Set page = Worksheets("indexes")
CheckBox1.Value = True
CheckBox4.Value = True
CheckBox3.Value = True
Call CheckBox1_Click
Call CheckBox3_Click
Call CheckBox4_Click
Call Label19_Click
Call Label18_Click

End Sub