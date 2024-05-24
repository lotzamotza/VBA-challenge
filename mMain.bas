Attribute VB_Name = "mMain"
Sub VBAChallenge()
    ' Workbook array variable is predefined because the number
    ' of workbooks to iterate over is known
    Dim wbStocks(0 To 1) As Workbook
    
    ' Worksheet array variable is dynamic because different number
    ' of sheets in each workbook
    Dim wksStocks() As Worksheet
    
    ' Same variables will hold the range for each worksheet, taking on
    ' new location for each sheet depending on the number of rows
    Dim rngStocks As Range
    Dim rngOutput As Range
    Dim rngSolution As Range
    
    ' For conditional formatting in output range
    Dim fcChange(0 To 1) As FormatCondition
    
    ' Variables for storing calculation results
    Dim sFilePath() As String
    Dim sTickerSymbol As String
    Dim sMaxPercentageIncreaseTicker As String
    Dim sMaxPercentageDecreaseTicker As String
    Dim sMaxTotalVolumeTicker As String
    Dim dOpenPrice As Double
    Dim dClosePrice As Double
    Dim dQuarterlyChange As Double
    Dim dPercentageChange As Double
    Dim dTotalVolume As Double
    Dim dMaxPercentageIncrease As Double
    Dim dMaxPercentageDecrease As Double
    Dim dMaxTotalVolume As Double
    Dim lNumRows As Long
    
    ' Loop variables & the independent iteration variable, x
    Dim i As Long, j As Long
    Dim k As Long, x As Long
    
    ' If an error occurs while the process is running it will
    ' direct control to the label ErrHandler and exit with a message
    On Error GoTo ErrHandler
    
    ' Stops events, turns calculation mode to manual, stops alerts, &
    ' stops screen updating to potentially run routine faster
    Call EventStop
    
    ' Allow user to select workbooks to operate with
    sFilePath = ExcelFileSelection
    
    If bError Then GoTo ErrHandler
    
    ' Set workbook variables
    Set wbStocks(0) = Workbooks.Open(sFilePath(0))
    Set wbStocks(1) = Workbooks.Open(sFilePath(1))
    
    ' Main outer loop is for iterating through both workbooks
    For i = 0 To 1
        
        ' Set dimensions - lower bound is zero, upper bound is corresponding
        ' to the number of sheets in currently iterated workbook
        ReDim wksStocks(0 To wbStocks(i).Sheets.Count)
        
        ' First inner loop iterates over each sheet in currently iterated workbook
        ' It also sets appropriate range variables and creates column headers
        For j = 0 To UBound(wksStocks) - 1
            
            ' Setting worksheet and its range variables
            Set wksStocks(j) = wbStocks(i).Sheets(j + 1)
            Set rngStocks = wksStocks(j).Range("A1:G1")
            Set rngStocks = wksStocks(j).Range(rngStocks, rngStocks.Offset(1000000).End(xlUp))
            Set rngOutput = wksStocks(j).Range("I1")
            Set rngSolution = wksStocks(j).Range("N1:P4")
            
            ' If script has run before, clear previous output data...
            rngOutput.CurrentRegion.ClearContents
            rngSolution.ClearContents
            
            ' The initial values were stored as text, which requires correction to
            ' perform calculations.
            If rngStocks.Cells(2, 1).Errors(xlNumberAsText).Value Then Call ConvertFromText(rngStocks)
            
            ' Set column & row headers for currently iterated sheet
            rngOutput.Cells(1, 1).Value = "Ticker"
            rngOutput.Cells(1, 2).Value = "Quarterly Change"
            rngOutput.Cells(1, 3).Value = "Percentage Change"
            rngOutput.Cells(1, 4).Value = "Volume"
            rngSolution.Cells(2, 1).Value = "Greatest % Increase"
            rngSolution.Cells(3, 1).Value = "Greatest % Decrease"
            rngSolution.Cells(4, 1).Value = "Greatest Total Volume"
            rngSolution.Cells(1, 2).Value = "Ticker"
            rngSolution.Cells(1, 3).Value = "Value"
            
            ' Due to nature of next loop, we need to set initial values so that
            ' the condition is not immediately tripped by an unlike sTickerSymbol
            ' Also get number of rows in rngStocks to iterate over
            sTickerSymbol = rngStocks.Cells(2, 1).Value
            dOpenPrice = rngStocks.Cells(2, 3).Value
            lNumRows = rngStocks.Rows.Count
        
            ' The counter variable x must be set to 0 to avoid output
            ' to incorrect cell since it is not a loop variable
            x = 0
            
            ' In this second inner loop we iterate over each row in rngStocks
            ' If the stock ticker changes, then we know that we can calculate
            ' certain values and store in variables
            For k = 0 To lNumRows - 2
                
                ' Check if currently iterated row ticker symbol matches symbol stored
                ' in sTickerSymbol variable
                If sTickerSymbol <> rngStocks.Cells(k + 2, 1).Value Then
                    
                    ' If true, then calculate last stock items
                    ' then output variable values to rngOutput cell, &
                    ' increase the value of the x counter variable, &
                    ' finally, switch to next stock ticker
                    dClosePrice = rngStocks.Cells(k + 1, 6).Value
                    dQuarterlyChange = WorksheetFunction.Round(dClosePrice - dOpenPrice, 2)
                    dPercentageChange = WorksheetFunction.Round(dQuarterlyChange / dClosePrice, 4)
                    
                    rngOutput.Cells(x + 2, 1).Value = sTickerSymbol
                    rngOutput.Cells(x + 2, 2).Value = dQuarterlyChange
                    rngOutput.Cells(x + 2, 3).Value = dPercentageChange
                    rngOutput.Cells(x + 2, 4).Value = dTotalVolume
                    
                    x = x + 1
                    
                    sTickerSymbol = rngStocks.Cells(k + 2, 1).Value
                    dOpenPrice = rngStocks.Cells(k + 2, 3).Value
                    dTotalVolume = rngStocks.Cells(k + 2, 7).Value
                Else
                    ' If false, continue cumulative sum for total volume
                    dTotalVolume = dTotalVolume + rngStocks.Cells(k + 2, 7).Value
                End If
                
            Next k
            
            ' Before moving to next sheet, set some formatting for rngOutput
            ' This requires changing the rngOutput variables address
            Set rngOutput = Intersect(rngOutput.CurrentRegion, wksStocks(j).Rows("2:100000"))
            
            rngOutput.Columns(1).NumberFormat = "@"
            rngOutput.Columns(2).NumberFormat = "#,##0.00"
            rngOutput.Columns(3).NumberFormat = "0.0#%"
            rngOutput.Columns(4).NumberFormat = "#,##0"
            rngSolution.Columns(2).NumberFormat = "@"
            rngSolution.Cells(2, 3).NumberFormat = "0.0#%"
            rngSolution.Cells(3, 3).NumberFormat = "0.0#%"
            rngSolution.Cells(4, 3).NumberFormat = "#,##0"
            
            ' In case of multiple runs, delete current formatting conditions
            rngOutput.FormatConditions.Delete
            
            ' Conditional formatting variable set for second column in output range using
            ' the criteria that the value is either greater or less than zero to determine
            ' which color to apply
            Set fcChange(0) = rngOutput.Columns(2).FormatConditions.Add(xlCellValue, xlGreater, "=0")
            Set fcChange(1) = rngOutput.Columns(2).FormatConditions.Add(xlCellValue, xlLess, "=0")
            
            ' Applies color choices of green and red shades
            fcChange(0).Interior.Color = 5287936
            fcChange(1).Interior.Color = 9737946
            
            ' Next, determine the largest % increase/decrease & volume & their ticker symbols
            dMaxPercentageIncrease = WorksheetFunction.Max(rngOutput.Columns(3))
            dMaxPercentageDecrease = WorksheetFunction.Min(rngOutput.Columns(3))
            dMaxTotalVolume = WorksheetFunction.Max(rngOutput.Columns(4))

            sMaxPercentageIncreaseTicker = WorksheetFunction.Index(rngOutput, WorksheetFunction. _
                                           Match(dMaxPercentageIncrease, rngOutput.Columns(3), 0), 1)
            sMaxPercentageDecreaseTicker = WorksheetFunction.Index(rngOutput, WorksheetFunction. _
                                           Match(dMaxPercentageDecrease, rngOutput.Columns(3), 0), 1)
            sMaxTotalVolumeTicker = WorksheetFunction.Index(rngOutput, WorksheetFunction. _
                                           Match(dMaxTotalVolume, rngOutput.Columns(4), 0), 1)

            rngSolution.Cells(2, 2).Value = sMaxPercentageIncreaseTicker
            rngSolution.Cells(3, 2).Value = sMaxPercentageDecreaseTicker
            rngSolution.Cells(4, 2).Value = sMaxTotalVolumeTicker
            rngSolution.Cells(2, 3).Value = dMaxPercentageIncrease
            rngSolution.Cells(3, 3).Value = dMaxPercentageDecrease
            rngSolution.Cells(4, 3).Value = dMaxTotalVolume
            
            ' Autofits the output columns
            rngOutput.Columns.AutoFit
            rngSolution.Columns.AutoFit
        Next j
    
    Next i
    
    MsgBox "The process successfully completed!", vbOKOnly, "Attention"
    Exit Sub
    
ErrHandler:
    
    MsgBox "An error occurred, so the process has been terminated.", vbCritical, "Attention"
    
End Sub

Public Sub ConvertFromText(rng As Range)
    Dim c As Range
    Dim i As Long
    
    rng.NumberFormat = "General"
    
    For Each c In rng.Cells
        If Left(c.Value, 1) = "'" Then c.Value = Right(c.Value, Len(c.Value) - 1)
    
        c.Value = c.Value
    Next c
End Sub
