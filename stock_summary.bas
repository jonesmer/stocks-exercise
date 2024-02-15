Attribute VB_Name = "Module1"
'see README for assumptions, pseudocode, and acknowledgements.

Sub stockTotals():
'for use with "Multiple_year_stock_data.xlsx"
    
    'variables
    Dim ticker As String
    Dim openVal, changeNum, changePct, totalVol As Double
    Dim lastRow, lastCol, i, wsheet, summRow As Integer
    Dim rundownMaxPct, rundownMinPct, RundownTotVol As Double
    Dim rdMaxPctTicker, rdMinPctTicker, rdMaxVolTicker As String
    
    For wsheet = 1 To ThisWorkbook.Sheets.Count 'loop for whole sheet
        Worksheets(wsheet).Activate
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
        
        'this data should be sorted already, but in case it isn't...
        'sort raw data by ticker and date
        Range(Cells(1, 1), Cells(lastRow, lastCol)).Sort key1:=Range("A2"), key2:=Range("B2")
                
        'set up summary table headers
        Cells(1, lastCol + 2).Value = "Ticker"
        Cells(1, lastCol + 3).Value = "Yearly Change ($)"
        Cells(1, lastCol + 4).Value = "Percent Change"
        Cells(1, lastCol + 5).Value = "Total Stock Volume"
        summRow = 2
        
        'assign variables from row 2 of raw data
        ticker = Cells(2, 1).Value
        openVal = Cells(2, 3).Value
        changeNum = Cells(2, 6).Value - openVal
        If openVal <> 0 Then 'can't divide by 0
            changePct = changeNum / openVal
        Else: changePct = 0
        End If
        totalVol = Cells(2, 7).Value
        
        'initially setting rundown values to what's in row 2, for comparison later
        rundownMaxPct = changePct
        rundownMinPct = changePct
        RundownTotVol = totalVol
        rdMaxPctTicker = ticker
        rdMinPctTicker = ticker
        rdMaxVolTicker = ticker
                
        'loop through the rows
        For i = 3 To lastRow
            'is ticker(current row) different from ticker(saved)?
            If Cells(i, 1).Value <> ticker Then
            
                'print saved variables to curr row of summary
                Cells(summRow, lastCol + 2).Value = ticker
                Cells(summRow, lastCol + 3).Value = changeNum
                If changeNum < 0 Then 'conditional formatting for Yearly Change
                    Cells(summRow, lastCol + 3).Interior.ColorIndex = 3
                Else: Cells(summRow, lastCol + 3).Interior.ColorIndex = 4
                End If
                Cells(summRow, lastCol + 4).Value = changePct
                Cells(summRow, lastCol + 4).NumberFormat = "0.00%"
                If changePct < 0 Then 'conditional formatting for Pct Change
                    Cells(summRow, lastCol + 4).Interior.ColorIndex = 3
                Else: Cells(summRow, lastCol + 4).Interior.ColorIndex = 4
                End If
                Cells(summRow, lastCol + 5).Value = totalVol
                
                'check saved vars for rundown min/max
                '(if new max, save value and ticker name)
                If changePct > rundownMaxPct Then
                    rundownMaxPct = changePct
                    rdMaxPctTicker = ticker
                End If
                If changePct < rundownMinPct Then
                    rundownMinPct = changePct
                    rdMinPctTicker = ticker
                End If
                If totalVol > RundownTotVol Then
                    RundownTotVol = totalVol
                    rdMaxVolTicker = ticker
                End If
                
                'go to next row of summary
                summRow = summRow + 1
                
                'calculate variables for i
                ticker = Cells(i, 1).Value
                openVal = Cells(i, 3).Value
                changeNum = Cells(i, 6).Value - openVal
                If openVal <> 0 Then
                    changePct = changeNum / openVal
                Else: changePct = 0
                End If
                totalVol = Cells(i, 7).Value
            Else:
                'calculate variables; ticker and openval unchanged
                changeNum = Cells(i, 6).Value - openVal
                If openVal <> 0 Then
                    changePct = changeNum / openVal
                Else: changePct = 0
                End If
                totalVol = totalVol + Cells(i, 7).Value
            End If
        Next i
        
        '---------
        'this block is just for the very last row
        '-same as block in For loop, but since the info isn't printed until
        'the next row, we need to print the last row outside the loop and
        'check the final numbers for the last ticker against
        'max/minpct and maxvol
        
        'print final row of summ
        Cells(summRow, lastCol + 2).Value = ticker
        Cells(summRow, lastCol + 3).Value = changeNum
        If changeNum < 0 Then
            'conditional formatting for Yearly Change
            Cells(summRow, lastCol + 3).Interior.ColorIndex = 3
            Else: Cells(summRow, lastCol + 3).Interior.ColorIndex = 4
        End If
        Cells(summRow, lastCol + 4).Value = changePct
        Cells(summRow, lastCol + 4).NumberFormat = "0.00%"
        If changePct < 0 Then 'conditional formatting for Pct Change
            Cells(summRow, lastCol + 4).Interior.ColorIndex = 3
        Else: Cells(summRow, lastCol + 4).Interior.ColorIndex = 4
        End If
        Cells(summRow, lastCol + 5).Value = totalVol
        
        'check saved vars for rundown min/max
        '(if new max, save value and ticker name)
        If changePct > rundownMaxPct Then
            rundownMaxPct = changePct
            rdMaxPctTicker = ticker
        End If
        If changePct < rundownMinPct Then
            rundownMinPct = changePct
            rdMinPctTicker = ticker
        End If
        If totalVol > RundownTotVol Then
            RundownTotVol = totalVol
            rdMaxVolTicker = ticker
        End If
        '---------
        
        'set up rundown table (row/column titles)
        Cells(1, lastCol + 9).Value = "Ticker"
        Cells(1, lastCol + 10).Value = "Value"
        Cells(2, lastCol + 8).Value = "Greatest % Increase"
        Cells(3, lastCol + 8).Value = "Greatest % Decrease"
        Cells(4, lastCol + 8).Value = "Greatest Total Volume"

        'print maxPctInc and corresponding ticker
        Cells(2, lastCol + 9).Value = rdMaxPctTicker
        Cells(2, lastCol + 10).Value = rundownMaxPct
        Cells(2, lastCol + 10).NumberFormat = "0.00%"
        
        'print minPctInc and corresponding ticker
        Cells(3, lastCol + 9).Value = rdMinPctTicker
        Cells(3, lastCol + 10).Value = rundownMinPct
        Cells(3, lastCol + 10).NumberFormat = "0.00%"
        
        'print maxTotVol and corresponding ticker
        Cells(4, lastCol + 9).Value = rdMaxVolTicker
        Cells(4, lastCol + 10).Value = RundownTotVol
        Cells(4, lastCol + 10).NumberFormat = "##0.0E+0"
        
        'formatting so we can see full headers! :)
        Range("A:Q").Columns.AutoFit
        Range("A1").Activate

    Next wsheet
    Worksheets(1).Activate
        
End Sub

'clear out the formatting (created for debugging purposes)
Sub clearData():

    For wsheet = 1 To ThisWorkbook.Sheets.Count
        Worksheets(wsheet).Activate
        Range("I:Q").Clear
        Range("A:Q").Columns.ColumnWidth = 8.43
        Range("A1").Activate
    Next wsheet
    Worksheets(1).Activate
    
End Sub

