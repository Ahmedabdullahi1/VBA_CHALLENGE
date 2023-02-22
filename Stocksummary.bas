Attribute VB_Name = "Module1"
Sub SummarizationofStock()
    ' Grab each sheet in the book
    For Each sht In Worksheets
        ' Create a headings array
        Dim Headings() As Variant
        Headings = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
        HeadingsLength = (UBound(Headings) - LBound(Headings) + 1)
        
        'Get the length of the initial columns
        rowLength = sht.Range("A1", sht.Range("A1").End(xlToRight)).Columns.Count
        
        ' Apply headings array to create the 4 headings by iterating through Headings array
        Dim HeadingsIndex As Integer
        HeadingsIndex = 0
        For r = (rowLength + 2) To (HeadingsLength + rowLength + 1)
            sht.Cells(1, r).Value = Headings(HeadingsIndex)
            HeadingsIndex = HeadingsIndex + 1
        Next r
         
         
        ' Get the length of the entire column
        ColLength = sht.Range("A1", sht.Range("A1").End(xlDown)).Rows.Count
        
        
        Dim TickerName As String
        
        ' Set initial total value for each ticker
        TotalVol = 0
        
        ' Set initial row for the outputs from the second row
        WorkingRow = 2
        
        ' Need to assign the value of year start for each ticker
        YearStart = sht.Range("C2").Value
        
        ' Code block to iterate through each row
        For i = 2 To ColLength
        
            ' If the next row's value is not the same as the current row's
            If sht.Cells(i + 1, 1).Value <> sht.Cells(i, 1).Value Then
                TickerName = sht.Cells(i, 1).Value
                
                ' Before each loop starts, the closing value of the ticker's year needs to be assigned
                YearEnd = sht.Cells(i, 6).Value
                
                TotalVol = TotalVol + sht.Cells(i, 7).Value
                sht.Range("I" & WorkingRow).Value = TickerName
                sht.Range("L" & WorkingRow).Value = TotalVol
                sht.Range("J" & WorkingRow).Value = YearEnd - YearStart
                sht.Range("K" & WorkingRow).Value = Round(((YearEnd - YearStart) / YearStart) * 100, 2) & "%"
                
                If sht.Range("J" & WorkingRow).Value >= 0 Then
                    sht.Range("J" & WorkingRow).Interior.ColorIndex = 4
                Else
                    sht.Range("J" & WorkingRow).Interior.ColorIndex = 3
                End If
                
                
                ' Need to increase the row at the end of the iteration
                WorkingRow = WorkingRow + 1
                
                ' Need to reset the value of the next ticker's open value of the year
                YearStart = sht.Cells(i + 1, 3).Value
                
                TotalVol = 0
                
            ' If the next row's value is the same as the current row's
            Else
                TotalVol = TotalVol + sht.Cells(i, 7).Value
            End If
                
        Next i
        
        ' Set the new Headings Manually
        sht.Range("O2").Value = "Greatest % increase"
        sht.Range("O3").Value = "Greatest % decrease"
        sht.Range("O4").Value = "Greatest Volume increase"
        sht.Range("P1").Value = "Ticker"
        sht.Range("Q1").Value = "Value"
        
        
        'Dim NewColLength As Integer
        
        NewColLength = sht.Range("I1", sht.Range("I1").End(xlDown)).Rows.Count
        
        MaxPercIncrease = (WorksheetFunction.Max(sht.Range("K2:K" & NewColLength)))
        MaxPercDecrease = (WorksheetFunction.Min(sht.Range("K2:K" & NewColLength)))
        MaxVolume = (WorksheetFunction.Max(sht.Range("L2:L" & NewColLength)))
        
        
        sht.Range("Q2").Value = MaxPercIncrease * 100 & "%"
        sht.Range("Q3").Value = MaxPercDecrease * 100 & "%"
        sht.Range("Q4").Value = MaxVolume
        
        ' Iterate through the combined results and insert values into fields
        For j = 2 To NewColLength
            If sht.Cells(j, 11).Value = MaxPercIncrease Then
                sht.Range("P2").Value = sht.Cells(j, 11).Offset(0, -2)
            End If
            If sht.Cells(j, 11).Value = MaxPercDecrease Then
                sht.Range("P3").Value = sht.Cells(j, 11).Offset(0, -2)
            End If
            If sht.Cells(j, 12).Value = MaxVolume Then
                sht.Range("P4").Value = sht.Cells(j, 12).Offset(0, -3)
            End If
        
        Next j
        
        
        ' Adjust all columns to autofit contents
        sht.Range("I:Q").Columns.AutoFit
        
    Next sht
    
End Sub

