Attribute VB_Name = "Module1"
Sub EverySheet()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call SMAnalysis

    Next
    Application.ScreenUpdating = True
End Sub

Sub SMAnalysis()
Dim sht As Worksheet

Dim LastColumn As Long
Dim currentticker As String
Dim pChange As Double
Dim YearlyChange As Double
Dim runClosing As Double
Dim runOpening As Double
Dim outputrow As Long
Dim LastRow As Long

Dim i As Long
Dim Vol As Double
Dim Largest As Double
Dim Smallest As Double
Dim LargestVol As Double
Dim SmallTicker As String
Dim LargeTicker As String

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

'determine last row
LastRow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row


'What is the last column
LastColumn = sht.Range("A1").CurrentRegion.Columns.Count

'create new headings
Cells(1, LastColumn + 2) = "Ticker"
Cells(1, LastColumn + 3) = "Yearly Change"
Cells(1, LastColumn + 4) = "Percent Change"
Cells(1, LastColumn + 5) = "Total Stock Volume"
Cells(1, LastColumn + 8) = "Ticker"
Cells(1, LastColumn + 9) = "Value"
Cells(2, LastColumn + 7) = "Greatest % Increase"
Cells(3, LastColumn + 7) = "Greatest % Decrease"
Cells(4, LastColumn + 7) = "Greatest Total Volume"

  Column = 1

  ' Loop through rows in the column
    For i = 2 To LastRow
            
'store current row in running totals
    runClosing = Cells(i, 6) + runClosing
    runOpening = Cells(i, 3) + runOpening
    currentticker = Cells(i, 1)
    Vol = Cells(i, 7) + Vol
    
   
' Searches for when the value of the next cell is different than that of the current cell
    If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
    
'next ticker is different so output cells
        Cells(outputrow, LastColumn + 2) = currentticker
        YearlyChange = runClosing - runOpening
        If runOpening <> 0 Then
            pChange = YearlyChange / runOpening
            Else: pChange = 0
        End If
        Cells(outputrow, LastColumn + 3) = YearlyChange
        Cells(outputrow, LastColumn + 4) = pChange
        Cells(outputrow, LastColumn + 5) = Vol
        
'is this the first output, this prevents an error i got in testing
' when the total was negative and it struggled to correctly reflect the 'greatest' changes
        If outputrow = 2 Then
                LargeVolTicker = currentticker
                LargestVol = Vol
                LargeTicker = currentticker
                Largest = pChange
                SmallTicker = currentticker
                Smallest = pChange
                Else
                
    'this is not the first output so check existing entry
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
       
' conditional formatting
        If YearlyChange < 0 Then
            Cells(outputrow, LastColumn + 3).Interior.ColorIndex = 3
            Else
            Cells(outputrow, LastColumn + 3).Interior.ColorIndex = 4
        End If
                If pChange < 0 Then
            Cells(outputrow, LastColumn + 4).Interior.ColorIndex = 3
            Else
            Cells(outputrow, LastColumn + 4).Interior.ColorIndex = 4
        End If
        
'then reset ready for the next ticker
        pChange = 0
        runClosing = 0
        runOpening = 0
        Vol = 0
        outputrow = outputrow + 1
   
    End If
    

Next i

'end of loop, now to output the final summary data

Cells(2, LastColumn + 8) = LargeTicker
Cells(2, LastColumn + 9) = Largest
Cells(3, LastColumn + 8) = SmallTicker
Cells(3, LastColumn + 9) = Smallest
Cells(4, LastColumn + 8) = LargeVolTicker
Cells(4, LastColumn + 9) = LargestVol

'format columns
sht.Columns(LastColumn + 3).NumberFormat = "0.00"
sht.Columns(LastColumn + 3).Cells.HorizontalAlignment = xlHAlignRight
sht.Columns(LastColumn + 4).NumberFormat = "0.00%"
sht.Columns(LastColumn + 4).Cells.HorizontalAlignment = xlHAlignRight

''format summary fields
Cells(2, LastColumn + 9).NumberFormat = "0.00%"
Cells(3, LastColumn + 9).NumberFormat = "0.00%"
Cells(4, LastColumn + 9).NumberFormat = "0"
sht.Columns(LastColumn + 9).Cells.HorizontalAlignment = xlHAlignRight
sht.Columns(LastColumn + 8).Cells.HorizontalAlignment = xlHAlignLeft

End Sub


