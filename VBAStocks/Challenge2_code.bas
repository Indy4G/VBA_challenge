Attribute VB_Name = "Module1"
Sub stocks():
'Note: this program assumes that the data on each sheet is in alphabetical order by ticker
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet 'Will set my active sheet back after iterating through. *See second to last line of code

For Each ws In Sheets
    'Cleans all rows affected (populated) by the program to allow for changes and a rerun
    ws.Columns(9).ClearContents
    ws.Columns(10).ClearContents
    ws.Columns(10).Interior.Color = xlNone
    ws.Columns(11).ClearContents
    ws.Columns(12).ClearContents

    ws.Columns(15).ClearContents
    ws.Columns(16).ClearContents
    ws.Columns(17).ClearContents
    '-----------------------------------------------------------------------
    'Headers and Labels
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    '------------------------------------------------------------------------
    'Defining all variables
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
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
    '------------------------------------------------------------------------
    'Actually doing stuff
    For i = 2 To lastrow
        
        totalvol = totalvol + ws.Cells(i, 7).Value
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value
        
            ws.Cells(rowcount, 10).Value = (ws.Cells(i, 6).Value - ws.Cells(firstopen, 3).Value)
            If ws.Cells(rowcount, 10).Value < 0 Then
                ws.Cells(rowcount, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(rowcount, 10).Interior.ColorIndex = 4
            End If
            
            If ws.Cells(firstopen, 3).Value = 0 Then
                ws.Cells(rowcount, 11).Value = Format(0, "Percent")
            Else
                ws.Cells(rowcount, 11).Value = Format((ws.Cells(rowcount, 10).Value / ws.Cells(firstopen, 3).Value), "Percent")
            End If
            
            ws.Cells(rowcount, 12).Value = totalvol
        
            rowcount = rowcount + 1
            firstopen = (i + 1)
            totalvol = 0
        
        End If
        
    Next i
    
    For j = 2 To rowcount
    
        If ws.Cells(j, 11).Value > max Then
            max = ws.Cells(j, 11).Value
            maxticker = ws.Cells(j, 9).Value
        End If
        
        If ws.Cells(j, 11).Value < min Then
            min = ws.Cells(j, 11).Value
            minticker = ws.Cells(j, 9).Value
        End If
        
        If ws.Cells(j, 12).Value > mostvol Then
            mostvol = ws.Cells(j, 12).Value
            mostvolticker = ws.Cells(j, 9).Value
        End If
        
    Next j
    
    ws.Range("P2").Value = maxticker
    ws.Range("P3").Value = minticker
    ws.Range("P4").Value = mostvolticker
    ws.Range("Q2").Value = Format(max, "Percent")
    ws.Range("Q3").Value = Format(min, "Percent")
    ws.Range("Q4").Value = Format(mostvol, "Scientific")
    
Next ws

starting_ws.Activate 'Returns the user to original sheet

End Sub

