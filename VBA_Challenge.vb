Sub Repeater()
    ' This is the section that will repeat the lower Sub for each Sheet in the Workbook
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call stockAnalysis
    Next
    Application.ScreenUpdating = True
End Sub
Sub stockAnalysis()
' Defining all the terms
Dim i As Long
Dim j As Integer
Dim k As Long
Dim yearopen As Double
Dim yearclose As Double
Dim thechange As Double
Dim pmax As Double
Dim pmin As Double
Dim stockvol As Double
lastrow = Cells(Rows.Count, "A").End(xlUp).Row
Dim ColorRange As Range
Set ColorRange = Range("M2:M" & lastrow)
j = 0
max = 0
pmax = 0
pmin = 0
stockvol = 0
'For Loop to pull the data over for analysis
' Changing the column widths for bonus secion
Columns("P").ColumnWidth = 25
Columns("R").ColumnWidth = 15
For i = 2 To lastrow
    ' Finding where the break point exists in Column A, as in the change from A to AA
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        ' Taking the number of rows it just went through and removing one to find open value
        k = i - (k)
        ' Pulling the each unqiue Ticker Value
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        ' Pulling the Total Stock Values Over for each Ticker
        Range("J" & 2 + j).Value = Total
        ' Finding the closing value of each Stock
        Range("L" & 2 + j).Value = Cells(i, 6).Value
        'Finding the open value of each Stock
        Range("K" & 2 + j).Value = Cells(k, 3).Value
        ' Setting those cells to equal yearopen, yearclose
        yearopen = Range("K" & 2 + j).Value
        yearclose = Range("L" & 2 + j).Value
        ' Finding the change from open to close
        Range("M" & 2 + j).Value = yearclose - yearopen
        ' Finding the percentage of change and formatting as such
        thechange = Range("M" & 2 + j).Value

            If yearopen <> 0 Then
                Range("N" & 2 + j).Value = (thechange / yearopen)
                Range("N" & 2 + j).NumberFormat = "0.00%"
            End If
        ' Adding 1 to j and resetting Total and k counts
        j = j + 1
        Total = 0
        k = 0
    Else
        ' Finding the number of rows it has cycled through for each Ticker
        k = k + 1
        ' Adding up the Total Stock Vol for each Ticker
        Total = Total + Cells(i, 7)
    End If
  Next i

  'For loop that will return the three bonus criteria
  For i = 2 To lastrow
    ' Going through the list and finding if the current value is higher than the previous
    If Cells(i, 14) > pmax Then
    pmax = Cells(i, 14)
    Range("R2") = pmax
    Range("Q2") = Cells(i, 9)
    End If
    ' Going through the list and finding if the current value is lower than the previous
    If Cells(i, 14) < pmin Then
    pmin = Cells(i, 14)
    Range("R3") = pmin
    Range("Q3") = Cells(i, 9)
    End If
    ' Going through the list and finding the highest overall stock volume
    If Cells(i, 10) > stockvol Then
    stockvol = Cells(i, 10)
    Range("R4") = stockvol
    Range("Q4") = Cells(i, 9)
    End If
Next i
' Inserting headers for the columns and titles/bold for the bonus section
Range("I1") = "<ticker>"
Range("J1") = "<stockvol>"
Range("K1") = "<open>"
Range("L1") = "<close>"
Range("M1") = "<change>"
Range("N1") = "<percentage>"
Range("P2").Font.Bold = True
Range("P3").Font.Bold = True
Range("P4").Font.Bold = True
Range("P2") = "Greatest Increase"
Range("P3") = "Greatest Decrease"
Range("P4") = "Greatest Stock Vol"

' Change column colors here
ColorRange.FormatConditions.Delete
ColorRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0.0"
ColorRange.FormatConditions(1).Interior.Color = vbRed
ColorRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="0.0"
ColorRange.FormatConditions(2).Interior.Color = vbGreen






End Sub
