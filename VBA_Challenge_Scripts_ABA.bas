Attribute VB_Name = "Module2"
Sub Stocklooptest()
    ' Set variables that need to be tested
    Dim ticker, sym, nextSym, highestVolSym, highestPercSym, lowestPercSym As String
    Dim tickerCol, volCol, openCol, closeCol As Long
    Dim rowCount, colCount, ssRowCount, ssColCount, i, rowIter As Long
    Dim change, changeStart, changeEnd, percentChange, total, highestVol, highestPerc, lowestPerc As Double
    
    ' Add new sheet for script results & set results to Summary Tab
    Sheets.Add.Name = "Summary"
    Set ss = ThisWorkbook.Sheets("Summary")
    
    ssColCount = ss.UsedRange.Columns.Count
    ssRowCount = ss.UsedRange.Rows.Count
    ' Auto adjust cell borders
    ss.Cells.Borders.LineStyle = xlContinuous
    
    ' Start for loops to run through worksheet using parameters defined above
    For Each ws In Worksheets
        colCount = ws.UsedRange.Columns.Count
        rowCount = ws.UsedRange.Rows.Count
        ' Start for loops to start evaluattion of variables starting with tickers
            For i = 1 To colCount
              If InStr(ws.Cells(1, i), "ticker") = 0 Then
              Else
                tickerCol = i
                Exit For
              End If
            Next i
            
            For i = 1 To colCount
              If InStr(ws.Cells(1, i), "open") = 0 Then
              Else
                openCol = i
                Exit For
              End If
            Next i

            For i = 1 To colCount
              If InStr(ws.Cells(1, i), "close") = 0 Then
              Else
                closeCol = i
                Exit For
              End If
            Next i

            For i = 1 To colCount
              If InStr(ws.Cells(1, i), "vol") = 0 Then
              Else
                volCol = i
                Exit For
              End If
            Next i
    ' Start assessment loops of changes and outputs
        If volCol > 0 Then
            rowIter = 2
            changeStart = ws.Cells(rowIter, openCol)

      For i = 2 To rowCount

        sym = ws.Cells(i, tickerCol)
        nextSym = ws.Cells(i + 1, tickerCol)
        total = total + ws.Cells(i, volCol)
    ' Loop for percent changes
        If total > highestVol Then
          highestVol = total
          highestVolSym = ws.Cells(i, tickerCol)
        End If
    ' Loop for checking between ticker differences
        If sym <> nextSym Then
          ss.Cells(1, ssColCount) = "<ticker>"
          ss.Cells(1, ssColCount + 1) = "<change>"
          ss.Cells(1, ssColCount + 2) = "<percent_change>"
          ss.Cells(1, ssColCount + 3) = "<total_vol>"
    ' run output for ticker percent changes
          changeEnd = ws.Cells(i, closeCol)
          change = changeEnd - changeStart
    ' run loop to check for yearly percent change at begining v end of a year
          If changeEnd = 0 Then
            percentChange = 0
          Else
            percentChange = (changeEnd - changeStart) / changeEnd
    ' loop checking for highest percent changes
            If percentChange > highestPerc Then
              highestPerc = percentChange
              highestPercSym = ws.Cells(i, tickerCol)
            End If
    ' loop checking for lowest percent changes
            If percentChange < lowestPerc Then
              lowestPerc = percentChange
              lowestPercSym = ws.Cells(i, tickerCol)
            End If
          End If
    ' output of for loops
        ss.Cells(rowIter, ssColCount) = sym
        ss.Cells(rowIter, ssColCount + 1) = change
        ss.Cells(rowIter, ssColCount + 2) = percentChange
        ss.Cells(rowIter, ssColCount + 3) = total
    ' adjusting percent change cells to the right % formatting
        ss.Cells(rowIter, ssColCount + 2).NumberFormat = "0.00%"
    ' for loops adjusting cell colors due to % cell changes
         If ss.Cells(rowIter, ssColCount + 1) >= 0 Then
            ss.Cells(rowIter, ssColCount + 1).Interior.Color = RGB(0, 255, 0)
          Else
            ss.Cells(rowIter, ssColCount + 1).Interior.Color = RGB(255, 0, 0)
          End If
        rowIter = rowIter + 1
        sym = nextSym
        change = 0
        percentChange = 0
        total = 0
        changeStart = ws.Cells(i + 1, openCol)
        End If
      Next i
      
      ssColCount = ssColCount + 5

    End If
    ' Run bouns loops outputs
    For i = 1 To 4
      ss.Cells(1, ssColCount + 1) = "<ticker>"
      ss.Cells(1, ssColCount + 2) = "<value>"
      ss.Cells(2, ssColCount) = "Greatest % Increase"
      ss.Cells(2, ssColCount + 1) = highestPercSym
      ss.Cells(2, ssColCount + 2) = highestPerc
      ss.Cells(3, ssColCount) = "Greatest % Decrease"
      ss.Cells(3, ssColCount + 1) = lowestPercSym
      ss.Cells(3, ssColCount + 2) = lowestPerc
      ss.Cells(4, ssColCount) = "Greatest Total Volume"
      ss.Cells(4, ssColCount + 1) = highestVolSym
      ss.Cells(4, ssColCount + 2) = highestVol

      ss.Cells(2, ssColCount + 2).NumberFormat = "0.00%"
      ss.Cells(3, ssColCount + 2).NumberFormat = "0.00%"
    Next i

  Next ws
    
End Sub

