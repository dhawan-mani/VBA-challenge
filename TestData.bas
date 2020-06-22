Attribute VB_Name = "Module1"
Sub stocks()
    Dim totalRows As LongLong
    Dim rowCountForYearsChange As LongLong
    Dim cc_count As LongLong
    
    Dim firstRowForUniqueTicker As LongLong ' Storing first row number when we find unique ticker
    Dim totalVolume As LongLong
    Dim firstOpenForUniqueTicker As Double
    Dim j As LongLong

    cc_count = 2
    totalVolume = 0
    firstRowForUniqueTicker = 2
    
    ' Set headers for columns to be calculated
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Years Change"
    Cells(1, 11).Value = "Percent change"
    Cells(1, 12).Value = "Total Stock Volume"

    ' Total number of rows
    totalRows = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' In our sheet, the values are ordered with same ticker symbols arranged in order
    ' Therefore, we can assume that any change in ticker symbols means we are seeing a new Ticker symbol
    ' Thus, we can find "unique" ticker symbols with just 1 for loop
    For j = 2 To totalRows
       
        If Cells(j, 1).Value <> Cells(j + 1, 1).Value Then ' New ticker symbol found
            totalVolume = totalVolume + Cells(j, 7)
            
            Cells(cc_count, 9).Value = Cells(j, 1).Value ' Set the unique ticker symbol
            Cells(cc_count, 10).Value = Cells(j, 6).Value - Cells(firstRowForUniqueTicker, 3).Value ' Set % change
            firstOpenForUniqueTicker = Cells(firstRowForUniqueTicker, 3).Value
            ' formula for % change = YearsChange / First open for ticker
            If firstOpenForUniqueTicker > 0 Then
                Cells(cc_count, 11).Value = (Round((Cells(cc_count, 10).Value / firstOpenForUniqueTicker) * 100, 3))
            Else
                Cells(cc_count, 11).Value = 0
            End If
    '
            Cells(cc_count, 12).Value = totalVolume ' Set value of total stock volume
            cc_count = cc_count + 1
            firstRowForUniqueTicker = j + 1 ' We have found new unique ticker; reset row number for ticker here
            totalVolume = 0
        Else
            totalVolume = totalVolume + Cells(j, 7)
        End If
     Next j
     
    rowCountForYearsChange = Cells(Rows.Count, 10).End(xlUp).Row
    For j = 2 To rowCountForYearsChange
        If Cells(j, 10).Value < 0 Then
            Range("J" & j).Interior.ColorIndex = 3
        Else
            Range("J" & j).Interior.ColorIndex = 4
        End If
    Next j
    
    MsgBox ("The last row is :" + Str(totalRows))
End Sub

Sub Challenge()
    Dim m As LongLong
    Dim n As LongLong
    Dim row_MaxK As LongLong
    Dim row_MaxL As LongLong
    Dim row_MinK As LongLong
    Dim counter As LongLong
    
    
    ' Set headers for the calculated columns
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    
    
    m = Cells(Rows.Count, 11).End(xlUp).Row 'Calculate number of rows in Percent Change column
    n = Cells(Rows.Count, 12).End(xlUp).Row 'Calculate number of rows in Total stocks column
    
    ' Set the following values :
    ' 1. Max value of % change
    ' 2. Min value of % change
    ' 3. Max value of total stocks volume
    Cells(2, 17).Value = Application.WorksheetFunction.Max(Range("K2:K" & m))
    Cells(3, 17).Value = Application.WorksheetFunction.Min(Range("K2:K" & m))
    Cells(4, 17).Value = Application.WorksheetFunction.Max(Range("L2:L" & n))
    
    ' Find corresponding row to find Ticker symbols
    ' Find Ticker with max total stocks , column number 12
    For counter = 2 To n
        If Cells(counter, 12) = Cells(4, 17).Value Then
            row_MaxL = counter
            Exit For
        End If
    Next counter
    ' Find Ticker with max & min value in Percent Change
    For counter = 2 To m
        If Cells(counter, 11).Value = Cells(2, 17).Value Then
            row_MaxK = counter
        ElseIf Cells(counter, 11).Value = Cells(3, 17).Value Then
            row_MinK = counter
        End If
    Next counter
    
    Cells(2, 16).Value = Cells(row_MaxK, 9).Value
    Cells(3, 16).Value = Cells(row_MinK, 9).Value
    Cells(4, 16).Value = Cells(row_MaxL, 9).Value
    
    MsgBox ("Function over")
End Sub



