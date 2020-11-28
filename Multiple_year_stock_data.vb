# VBA-challenge

Sub EverySheet()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        xSh.Columns("K").NumberFormat = "0.00%"
        Call SMAnalysis

    Next
    Application.ScreenUpdating = True
End Sub
Sub SMAnalysis()
Dim i As Long
Dim sht As Worksheet
Dim currentticker As String
Dim pChange As Double
Dim YearlyChange As Double
Dim runClosing As Double
Dim runOpening As Double
Dim outputrow As Long
Dim LastRow As Long
Dim Vol As Double
Dim Largest As Double
Dim Smallest As Double
Dim LargestVol As Double
Dim SmallTicker As String
Dim LargeTicker As String
Dim LargeVolTicker As String

Set sht = ActiveSheet
outputrow = 2
pChange = 0
runOpening = 0
runClosing = 0
Vol = 0
LargestVol = 0
Largest = 0
Smallest = 0
currentticker = Cells(2, 1)
LastRow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row


  Column = 1

  ' Loop through rows in the column
    For i = 2 To LastRow
    
   'store current row
    runClosing = Cells(i, 6) + runClosing
    runOpening = Cells(i, 3) + runOpening
    currentticker = Cells(i, 1)
    Vol = Cells(i, 7) + Vol
    
' Searches for when the value of the next cell is different than that of the current cell
    If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
    
'next ticker is different so output cells
        Cells(outputrow, 9) = currentticker
        YearlyChange = runClosing - runOpening
        If runOpening <> 0 Then
            pChange = YearlyChange / runOpening
            Else: pChange = 0
        End If
        Cells(outputrow, 10) = YearlyChange
        Cells(outputrow, 11) = pChange
        Cells(outputrow, 12) = Vol
        'is this the first output
        If outputrow = 2 Then
                LargeVolTicker = currentticker
                LargestVol = Vol
                LargeTicker = currentticker
                Largest = pChange
                SmallTicker = currentticker
                Smallest = pChange
                Else
            If Vol > LargestVol Then
                LargeVolTicker = currentticker
                LargestVol = Vol
            End If
            If pChange >= Largest Then
                LargeTicker = currentticker
                Largest = pChange
            End If
            If pChange <= Smallest Then
                SmallTicker = currentticker
                Smallest = pChange
            End If
       End If

        ActiveSheet.Range("P2:P3").NumberFormat = "0.00%"
        
' conditional formatting
        If YearlyChange < 0 Then
            Cells(outputrow, 10).Interior.ColorIndex = 3
            Else
            Cells(outputrow, 10).Interior.ColorIndex = 4
        End If
                If pChange < 0 Then
            Cells(outputrow, 11).Interior.ColorIndex = 3
            Else
            Cells(outputrow, 11).Interior.ColorIndex = 4
        End If
        
'then reset ready for the next ticker
        pChange = 0
        runClosing = 0
        runOpening = 0
        Vol = 0
        outputrow = outputrow + 1
   
    End If
    

Next i

sht.Columns("K").NumberFormat = "0.00%"

Cells(2, 15) = LargeTicker
Cells(2, 16) = Largest
Cells(3, 15) = SmallTicker
Cells(3, 16) = Smallest
Cells(4, 15) = LargeVolTicker
Cells(4, 16) = LargestVol

End Sub
