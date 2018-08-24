Attribute VB_Name = "Module1"
Sub Analyze()
    
'Code to cycle through each worksheet
Dim xSh As Worksheet
For Each xSh In Worksheets
    xSh.Select
    

'Declare variables and populate with initial values
    'Note: there are more lines to loop through than can be handled by integers, so Long or longer is required
    Dim i As Long
    i = 2
    
    Dim tickernumber As Integer
    tickernumber = 2
    
    'Note: the total volume sold for some tickers got too large for even Long to handle, so LongLong or longer is required
    Dim CurrentVolume As LongLong
    
    Dim CurrentStr As String
    CurrentStr = Cells(2, 1).Value
    
    Dim CurrentOpen As Double
    Dim CurrentClosed As Double
    
    'declaring variables for tracking record setting tickers
    Dim CrntHighPcntVal As Double
    Dim CrntLowPcntVal As Double
    Dim CrntHighVolVal As LongLong
    
    Dim CrntHighPcntTicker As Long
    Dim CrntLowPcntTicker As Long
    Dim CrntHighVolTicker As Long
    
    'Resetting record values so as to avoid records from rolling over from sheet to sheet
    CrntHighPcntVal = vbNullValue
    CrntLowPcntVal = vbNullValue
    CrntHighVolVal = vbNullValue
    
    
    'Populating labels for outputted data
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    
    
    'Using a DoWhile current cell is nonempty to avoid needing to specify how long the columns are individually
    Do While Cells(i, 1).Value <> ""
        
        'Take note of what ticker we are currently tallying results for
        CurrentOpen = Cells(i, 3).Value
        CurrentVolume = 0
        
        'Keep tallying results for the current ticker until it changes to the next ticker symbol
        Do While Cells(i, 1).Value = CurrentStr
            CurrentVolume = CurrentVolume + Cells(i, 7)
            i = i + 1
        Loop
        
        'Once we're out of that loop, our i value points to the next ticker symbol
        'so look one line back for the final closing value for the ticker for the year
        CurrentClosed = Cells(i - 1, 6)
        
        'Populate output table with information about this ticker symbol
        
        Cells(tickernumber, 9) = CurrentStr
        'change from end of year's closing cost to beginning of year's opening cost
        Cells(tickernumber, 10) = CurrentClosed - CurrentOpen
        'Color background of cell with green if yearly change is positive or red if negative.  No color if no change
        If Cells(tickernumber, 10) > 0 Then
            Cells(tickernumber, 10).Interior.Color = RGB(0, 200, 0)
        ElseIf Cells(tickernumber, 10) < 0 Then
            Cells(tickernumber, 10).Interior.Color = RGB(200, 0, 0)
        End If
        
        'Error catching for division by zero
        If CurrentOpen <> 0 Then
            Cells(tickernumber, 11) = Cells(tickernumber, 10) / CurrentOpen
            'format as percentage
            Cells(tickernumber, 11).NumberFormat = "0.00%"
        
        'compare records for percentages in here to avoid issues with comparing to error message
            If CrntHighPcntVal < Cells(tickernumber, 11) Then
                CrntHighPcntVal = Cells(tickernumber, 11)
                CrntHighPcntTicker = tickernumber
            End If
            If CrntLowPcntVal > Cells(tickernumber, 11) Then
                CrntLowPcntVal = Cells(tickernumber, 11)
                CrntLowPcntTicker = tickernumber
            End If
        
        
        Else
            Cells(tickernumber, 11) = "Error: Division by Zero"
        End If
        
        Cells(tickernumber, 12) = CurrentVolume
        
        
        'compare and update for volume
        If CrntHighVolVal < CurrentVolume Then
            CrntHighVolVal = CurrentVolume
            CrntHighVolTicker = tickernumber
        End If
        
        
        
        'setup for looking at next ticker symbol
        tickernumber = tickernumber + 1
        CurrentStr = Cells(i, 1).Value
        
        
    Loop
    
    Range("P2").Value = Cells(CrntHighPcntTicker, 9).Value
    Range("Q2").Value = CrntHighPcntVal
    Range("Q2").NumberFormat = "0.00%"
    Range("P3").Value = Cells(CrntLowPcntTicker, 9).Value
    Range("Q3").Value = CrntLowPcntVal
    Range("Q3").NumberFormat = "0.00%"
    Range("P4").Value = Cells(CrntHighVolTicker, 9).Value
    Range("Q4").Value = CrntHighVolVal
    
Next






End Sub
