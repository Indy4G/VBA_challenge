Attribute VB_Name = "Module1"
Sub stocks():
'Note: this program assumes that the data is in alphabetical order by ticker

Columns(9).ClearContents
Columns(10).ClearContents
Columns(10).Interior.Color = xlNone
Columns(11).ClearContents
Columns(12).ClearContents

Columns(15).ClearContents
Columns(16).ClearContents
Columns(17).ClearContents

'---------------------------------------------------------------

'Headers and Labels
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
    
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'---------------------------------------------------------------

'Defining all variables
Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Dim rowcount As Long
rowcount = 2
Dim firstopen As Long
firstopen = 2
Dim totalvol As Double 'I don't understand why, but Long gave me an Overflow Error and Double works
totalvol = 0
    
Dim maxticker As String
Dim minticker As String
Dim mostvolticker As String
    
Dim max As Double
max = 0
Dim min As Double
min = 0
Dim mostvol As Double
mostvol = 0

'----------------------------------------------------------------

'Actually doing stuff
For i = 2 To lastrow
        
    totalvol = totalvol + Cells(i, 7).Value
        
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        Cells(rowcount, 9).Value = Cells(i, 1).Value
        
        Cells(rowcount, 10).Value = (Cells(i, 6).Value - Cells(firstopen, 3).Value)
        
        If Cells(rowcount, 10).Value < 0 Then
            Cells(rowcount, 10).Interior.ColorIndex = 3
        Else
            Cells(rowcount, 10).Interior.ColorIndex = 4
        End If
            
        If Cells(firstopen, 3).Value = 0 Then
            Cells(rowcount, 11).Value = Format(0, "Percent")
        Else
            Cells(rowcount, 11).Value = Format((Cells(rowcount, 10).Value / Cells(firstopen, 3).Value), "Percent")
        End If
            
        Cells(rowcount, 12).Value = totalvol
        
        rowcount = rowcount + 1
        firstopen = (i + 1)
        totalvol = 0
        
    End If
        
Next i
    
For j = 2 To rowcount
    
    If Cells(j, 11).Value > max Then
            max = Cells(j, 11).Value
            maxticker = Cells(j, 9).Value
    End If
        
    If Cells(j, 11).Value < min Then
        min = Cells(j, 11).Value
        minticker = Cells(j, 9).Value
    End If
        
    If Cells(j, 12).Value > mostvol Then
        mostvol = Cells(j, 12).Value
        mostvolticker = Cells(j, 9).Value
    End If
        
Next j
    
Range("P2").Value = maxticker
Range("P3").Value = minticker
Range("P4").Value = mostvolticker
Range("Q2").Value = Format(max, "Percent")
Range("Q3").Value = Format(min, "Percent")
Range("Q4").Value = Format(mostvol, "Scientific")

End Sub
