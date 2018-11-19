Attribute VB_Name = "StockReports"
'Indices for Input Values
Const I_TICKER = 0
Const I_DATE = 1
Const I_OPEN = 2
Const I_HIGH = 3
Const I_LOW = 4
Const I_CLOSE = 5
Const I_VOL = 6

'Indices for Output Values
Const O_TICKER = 0
Const O_YEARLYCHG = 1
Const O_PERCENTCHG = 2
Const O_TOTALSTOCKVOL = 3

'Additional Indices/References
Const OFFSET_CALCTABLE = 8
Const OFFSET_WINNERTABLE = 14
Const LIST_ORIGIN = "A1"

Function CalcStockMetrics(ByVal AWS As Worksheet) As Boolean
'For the specified worksheet, calculate stock metrics
' and populate them on a table on the same worksheet.
'
'KEY ASSUMPTIONS:
'   * The list of stock information may or may not be sorted by ticker => Arrays will be needed to hold per-stick tallies
'   * There are no more than 5,000 unique stocks (i.e., unique Tickers) => Array sizes will be 5000 max
'   * There are no more than 1,000,000 entries on the list of stock data per worksheet
'   * Tallies are performed on a per-worksheet basis (i.e., entries do not need to be combined across worksheets)
'
'Arguments:
'   AWS: The worksheet which should contain stock information
'
'Approach:
'   1. Loop through each row in the Stock table
'   2. For each row, perform several actions for the stock indicated by the ticker:
'       a. Keep track of the <open> and <close> values if they are closest to the beginning of the year
'       b. Keep track of the <open> and <close> values if they are closest to the end of the year
'       c. Maintain a cumulative sum of the <vol> of the stock
'       d. Store this information (including the ticker) in one or more Calculation arrays,
'           with each unique ticker using the same unique index
'
'   3. Finish processing all of the rows in the Stock Table
'
'   4. For each entry in the Calculation array(s):
'       a. Determine the Yearly Change as: beginning of year <open> - end of year <close>
'       b. Determine the Percent Change as: Yearly Change / beginning of year <open>
'       c. Keep track of Winner Ticker and Percent Change if the Percent Change is greatest % increase seen so far
'       d. Keep track of Winner Ticker and Percent Change if the Percent Change is greatest % decrease seen so far
'       e. Keep track of Winner Ticker and Total Stock Volume if the Total Stock Volume is greatest seen so far
'
'   5. Create the Calculation Table by writing a header and the entries from the Calculation array(s)
'       onto the spreadsheet AWS
'
'   6. Create the Winner Table by writing a header and the Winnder entries onto the spreadsheet AWS


    'Select the top-left cell on the target worksheet
    '(the underlying functions assume that we're working on the active sheet)
    AWS.Activate
    Range("A1").Select
    
    
    'First, check a value to see if a list is likely populated on this sheet
    If (InStr(1, AWS.Range(LIST_ORIGIN).Value, "ticker", vbTextCompare) <= 0) Then
        'Sorry, can't find the heading value that includes "ticker" in the
        ' spot where the origin of the table is supposed to be - exit!
        CalcStockMetrics = False
        Exit Function
    End If
    
    'Create 1-dimensional arrays to store info needed for calculation
    'Note: Could probably simplify things by using a Dictionary object to
    '       store the calculation info w/ key = Ticker symbol.
    '       But, will stick with basic 1-dim arrays and a slow search algorithm instead...
    Dim c_ticker(5000) As String    'Array w/ Ticker symbols
    Dim c_open_date(5000)           'Array w/ Date of earliest Open
    Dim c_open_val(5000) As Double  'Array w/ Open Value on earliest Open
    Dim c_close_date(5000)          'Array w/ Date of earliest Open
    Dim c_close_val(5000) As Double 'Array w/ Open Value on earliest Open
    Dim c_tsv(5000) As Variant      'Array w/ Sum of all daily Volumes
        
    'Starting building the calculation arrays at index = 0
    i_max = 0
    
    'This could take a little while so let's turn off Screen Updating (to speed things up)
    retval = ScreenUpdating(False)
    
    'Blank out the status bar
    retval = StatusBar_Msg()
    
    'Get the calculation data needed from each row in the worksheet
    'Loop through all rows on this worksheet until the <ticker> columns is empty
    'Remember the assumption that there are no more than 1,000,000 stock data entries per worksheet
    For r = 1 To 1000000
        If (IsEmpty(Range(LIST_ORIGIN).Offset(r, I_TICKER).Value)) Then
            Exit For
        End If
    
        'Let's keep the ticker
        t_stock = Range(LIST_ORIGIN).Offset(r, I_TICKER).Value
        
        'Let's send a nice progress message to the Status Bar for every 1000 rows processed
        If (r Mod 1000 = 0) Then
        retval = StatusBar_Msg("Processing Worksheet: " _
                        & AWS.Name & ", Stock Input Row: " & r _
                        & " (Address: " & Range(LIST_ORIGIN).Offset(r, I_TICKER).Address _
                        & ") = Ticker: " & Range(LIST_ORIGIN).Offset(r, I_TICKER).Value)
        End If
        
        'Check to see if this Ticker is already in the Calculation arrays
        'If not, then add it to the Calculation arrays
        retval = FindTickerIndex(c_ticker, t_stock)
        If (retval >= 0) Then
        
            'Return value is the index of this ticker in the array
            i_stock = retval
            
            'Update the entries in the Calculation arrays for this stock
            
            'If the Date of this entry is less than the current open date, store the new date and associated open value
            If (Range(LIST_ORIGIN).Offset(r, I_DATE).Value < c_open_date(i_stock)) Then
                c_open_date(i_stock) = Range(LIST_ORIGIN).Offset(r, I_DATE).Value
                c_open_val(i_stock) = Range(LIST_ORIGIN).Offset(r, I_OPEN).Value
            End If
            
            If (Range(LIST_ORIGIN).Offset(r, I_DATE).Value > c_close_date(i_stock)) Then
                c_close_date(i_stock) = Range(LIST_ORIGIN).Offset(r, I_DATE).Value
                c_close_val(i_stock) = Range(LIST_ORIGIN).Offset(r, I_CLOSE).Value
            End If
            
            'Add the stick Volume to the running total
            'Note: Use conversion to Decimal CDec() to permit storage of very large numbers without overflow
            c_tsv(i_stock) = c_tsv(i_stock) + CDec(Range(LIST_ORIGIN).Offset(r, I_VOL).Value)

        Else
            'If return value <0 then this entry needs to be added to the array
            i_stock = i_max
            
            c_ticker(i_stock) = t_stock
            c_open_date(i_stock) = Range(LIST_ORIGIN).Offset(r, I_DATE).Value
            c_open_val(i_stock) = Range(LIST_ORIGIN).Offset(r, I_OPEN).Value
            c_close_date(i_stock) = Range(LIST_ORIGIN).Offset(r, I_DATE).Value
            c_close_val(i_stock) = Range(LIST_ORIGIN).Offset(r, I_CLOSE).Value
            c_tsv(i_stock) = Range(LIST_ORIGIN).Offset(r, I_VOL).Value
            
            'Prepare for the next stock entry into the Calc arrays
            i_max = i_max + 1
            
        End If
        
    Next
    
    'Now... process the data in the calculation arrays and create the Calculation table
    
    'Clear the Status Bar
    retval = StatusBar_Msg()
    
    'Winners
    Dim winner_gpi_t As String, winner_gpd_t As String, winner_gtv_t As String
    Dim winner_gpi_v As Double, winner_gpd_v As Double
    Dim winner_gtv_v As Variant
    
    'Set to Empty to indicate no winner yet selected
    winner_gpi_t = Empty
    winner_gpd_t = Empty
    winner_gtv_t = Empty
    
    Dim yearlychg As Double, percentchg As Double
        
    'Populate the Calculation Table header
    retval = PopCalcTableHeader(Range(LIST_ORIGIN).Offset(0, OFFSET_CALCTABLE).Address)
    
    'Loop through all of the entries in the Calculation array(s)
    'Note: The last entry in the array is at: (i_max-1)
    For i = LBound(c_ticker) To i_max - 1
        'A little safety valve: If the Ticker string is empty (""), then exit this for loop
        If (Len(c_ticker(i)) = 0) Then
            Exit For
        End If

        'Let's send a nice progress message to the Status Bar for every 10 elements processed
        If (i Mod 10 = 0) Then
        retval = StatusBar_Msg("Calculating Stock Metrics: " _
                        & AWS.Name & ", Calculation Array Index: " & i _
                        & " = Ticker: " & c_ticker(i))
        End If


        'Ticker: c_ticker(i)
        
        'Calculate the yearly change
        yearlychg = c_close_val(i) - c_open_val(i)
        
        'Calculate the percent change, but avoid for divide by 0
        'If the open value was 0, then set percent change to 0%
        If (c_open_val(i) <> 0) Then
            percentchg = yearlychg / c_open_val(i)
        Else
            
            percentchg = 0
        End If
        
        'Total Stock Volume = c_tsv(i)
        
        'Populate 1 row in the Calculation Table for this stock
        retval = PopCalcTableRow(Range(LIST_ORIGIN).Offset(i + 1, OFFSET_CALCTABLE).Address, c_ticker(i), yearlychg, percentchg, c_tsv(i))
        'Function PopCalcTableRow(ByVal AAddress, ATVal As String, AYCVal As Double, APCVal As Double, ATSVVal As Long) As Boolean
        
        'Keep track of Winners
        'Greatest % Increase
        If (Len(winner_gpi_t) = 0 Or (percentchg > winner_gpi_v)) Then
            winner_gpi_t = c_ticker(i)
            winner_gpi_v = percentchg
        End If
        
        'Greatest % Decrease
        If (Len(winner_gpd_t) = 0 Or (percentchg < winner_gpd_v)) Then
            winner_gpd_t = c_ticker(i)
            winner_gpd_v = percentchg
        End If
        
        'Greatest Total Volume
        If (Len(winner_gtv_t) = 0 Or (c_tsv(i) > winner_gtv_v)) Then
            winner_gtv_t = c_ticker(i)
            winner_gtv_v = c_tsv(i)
        End If
    Next
    
    'Populate the Winner table
    retval = PopWinnerTable(Range(LIST_ORIGIN).Offset(0, OFFSET_WINNERTABLE).Address, winner_gpi_t, winner_gpi_v, winner_gpd_t, winner_gpd_v, winner_gtv_t, winner_gtv_v)
    
    'AutoFit the key columns on this worksheet so everything looks nice and neat
    AWS.Columns("A:Q").AutoFit
    
    'All done!  Clean-up
    
    'Turn Screen Updating back on
    retval = ScreenUpdating(True)
    
    'Blank out the status bar
    retval = StatusBar_Msg()
    
    'Set a successful return value before exiting
    CalcStockMetrics = True
    
End Function

Function PopCalcTableHeader(ByVal AAddress) As Boolean
'Write header entries for 4 columns to a spreadsheet at the specified address AAddress
'Note: The address should include the sheet name!
'
'Arguments:
'   AAddress: Address to populate on the spreadsheet
'
'Headings
'   O_TICKER: "Ticker" => Format: "General", Left aligned
'   O_YEARLYCHG: "Yearly Change" => Format: "General", Left aligned
'   O_PERCENTCHG: "Percent Change" => Format: "General", Left aligned
'   O_TOTALSTOCKVOL: "Total Stock Volume" => Format: "General", Left aligned
'
'Return Value:
'   TRUE: Successful
'   FALSE: Error due to invalid arguments or other issue

    'Populate each of the values in order
    With Range(AAddress).Offset(0, O_TICKER)
        .Value = "Ticker"
        .HorizontalAlignment = xlHAlignLeft
        .NumberFormat = "General"
        .Interior.Color = xlNone
    End With
    
    With Range(AAddress).Offset(0, O_YEARLYCHG)
        .Value = "Yearly Change"
        .HorizontalAlignment = xlHAlignLeft
        .NumberFormat = "General"
        .Interior.Color = xlNone
    End With
    
    With Range(AAddress).Offset(0, O_PERCENTCHG)
        .Value = "Percent Change"
        .HorizontalAlignment = xlHAlignLeft
        .NumberFormat = "General"
        .Interior.Color = xlNone
    End With

    With Range(AAddress).Offset(0, O_TOTALSTOCKVOL)
        .Value = "Total Stock Volume"
        .HorizontalAlignment = xlHAlignLeft
        .NumberFormat = "General"
        .Interior.Color = xlNone
    End With
    
    'Assume things were successful!
    PopCalcTableHeader = True

End Function

Function PopCalcTableRow(ByVal AAddress, ByVal ATVal As String, ByVal AYCVal As Double, ByVal APCVal As Double, ByVal ATSVVal As Variant) As Boolean
'Write the specified 4 arguments to a spreadsheet at the specified address AAddress
'Note: The address should include the sheet name!
'
'Arguments
'   AAddress: Address to populate on the spreadsheet
'   O_TICKER: ATVal: Ticker value => Format: Left aligned text
'   O_YEARLYCHG: AYCVal: Yearly Change => Format: Fixed (9 decimal places), Right aligned,
'                               Shading (Positive Change: Green, Negative Change: Red, No Change: No Color)
'   O_PERCENTCHG: APCVal: Percent Change => Format: Precent (2 decimal places), Right aligned
'   O_TOTALSTOCKVOL: ATSVVal: Total Stock Volume => Format: General, Right aligned
'
'Return Value:
'   TRUE: Successful
'   FALSE: Error due to invalid arguments or other issue

    'Populate each of the values in order
    With Range(AAddress).Offset(0, O_TICKER)
        .Value = ATVal
        .HorizontalAlignment = xlHAlignLeft
        .NumberFormat = "General"
        .Interior.Color = xlNone        'REMINDER: Green = vbGreen, Red = 'vbRed, No Fill = xlNone
    End With
    
    With Range(AAddress).Offset(0, O_YEARLYCHG)
        .Value = AYCVal
        .HorizontalAlignment = xlHAlignRight
        .NumberFormat = "0.000000000"
        
        'Set the Fill color of this cell based upon the value of the Yearly Change
        If (AYCVal > 0) Then
            .Interior.Color = vbGreen
        ElseIf (AYCVal < 0) Then
            .Interior.Color = vbRed
        Else
            .Interior.Color = xlNone
        End If
        
    End With
    
    With Range(AAddress).Offset(0, O_PERCENTCHG)
        .Value = APCVal
        .HorizontalAlignment = xlHAlignRight
        .NumberFormat = "0.00%"
        .Interior.Color = xlNone
    End With

    With Range(AAddress).Offset(0, O_TOTALSTOCKVOL)
        .Value = ATSVVal
        .HorizontalAlignment = xlHAlignRight
        .NumberFormat = "General"
        .Interior.Color = xlNone
    End With
    
    'Assume things were successful!
    PopCalcTableRow = True

End Function

Function PopWinnerTable(ByVal AAddress, ByVal AGPI_T As String, ByVal AGPI_V As Double, ByVal AGPD_T As String, ByVal AGPD_V As Double, ByVal AGTV_T As String, ByVal AGTV_V As Variant) As Boolean
'Create the Winner Table on a spreadsheet at the specified address AAddress
'Note: The address should include the sheet name!
'
'Arguments
'   O_TICKER: ATVal: Ticker value => Format: Left aligned text
'   O_YEARLYCHG: AYCVal: Yearly Change => Format: Fixed (9 decimal places), Right aligned,
'                               Shading (Positive Change: Green, Negative Change: Red, No Change: No Color)
'   O_PERCENTCHG: APCVal: Percent Change => Format: Precent (2 decimal places), Right aligned
'   O_TOTALSTOCKVOL: ATSVVal: Total Stock Volume => Format: General, Right aligned
'
'   AAddress: Address to populate on the spreadsheet
'   AGPI_T and AGPI_V: The Ticker with Greatest Percent Increase and associated Percent Change value
'   AGPD_T and AGPD_V: The Ticker with Greatest Percent Decrease and associated Percent Change value
'   AGTV_T and AGTV_V: The Ticker with Greatest Total Stock Volume and associated Total Stock Volume value
'

'Return Value:
'   TRUE: Successful
'   FALSE: Error due to invalid arguments or other issue

    'Populate Headings
    Range(AAddress).Offset(0, 1).Value = "Ticker"
    Range(AAddress).Offset(0, 2).Value = "Value"
    
    'Populate each of the values in order
    Range(AAddress).Offset(1, 0).Value = "Greatest % Increase"
    With Range(AAddress).Offset(1, 1)
        .Value = AGPI_T
        .HorizontalAlignment = xlHAlignLeft
        .NumberFormat = "General"
        .Interior.Color = xlNone
    End With
    
    With Range(AAddress).Offset(1, 2)
        .Value = AGPI_V
        .HorizontalAlignment = xlHAlignRight
        .NumberFormat = "0.00%"
        .Interior.Color = xlNone
    End With

    'Populate each of the values in order
    Range(AAddress).Offset(2, 0).Value = "Greatest % Decrease"
    With Range(AAddress).Offset(2, 1)
        .Value = AGPD_T
        .HorizontalAlignment = xlHAlignLeft
        .NumberFormat = "General"
        .Interior.Color = xlNone
    End With
    
    With Range(AAddress).Offset(2, 2)
        .Value = AGPD_V
        .HorizontalAlignment = xlHAlignRight
        .NumberFormat = "0.00%"
        .Interior.Color = xlNone
    End With

    'Populate each of the values in order
    Range(AAddress).Offset(3, 0).Value = "Greatest Total Volume"
    With Range(AAddress).Offset(3, 1)
        .Value = AGTV_T
        .HorizontalAlignment = xlHAlignLeft
        .NumberFormat = "General"
        .Interior.Color = xlNone
    End With
    
    With Range(AAddress).Offset(3, 2)
        .Value = AGTV_V
        .HorizontalAlignment = xlHAlignRight
        .NumberFormat = "General"
        .Interior.Color = xlNone
    End With
    
    'Assume things were successful!
    PopWinnerTable = True

End Function
Function FindTickerIndex(ByRef AArrayName() As String, ByVal AMatchValue As String) As Long
'Loop through a specified array AArrayName to find the array index where the value matches the
' specified AMatchValue.  Make the comparison as a text/string comparison.
'
'Arguments:
'   AArrayName: A 1-dimensional array to be searched
'   AMatchValue: The value to be found in the array
'
'Return Value:
'   -1 => No match found
'   >=0 => Index of the array element matching AMatchValue

'Loop through all elements in the array
For i = LBound(AArrayName) To UBound(AArrayName)
    If (Not (IsNull(AArrayName(i))) And StrComp(AArrayName(i), AMatchValue, vbTextCompare) = 0) Then
        'Found a match - return the index at which it was found
        FindTickerIndex = i
        Exit Function
    End If
Next

'Ok, we we got this far then there was no match!
FindTickerIndex = -1

End Function



Sub ProcessAllWorksheets()
    For Each ws In Worksheets
        Debug.Print "Processing Worksheet Name: " & ws.Name
        retval = CalcStockMetrics(ws)
    Next
    
    'retval = PopCalcTableHeader("Sheet9!I2")
    'retval = PopCalcTableRow("Sheet9!I3", "JAB", 2.00001, 0.234, 987888374)
    'retval = PopWinnerTable("Sheet9!O2", "ABC", 2.3, "DEF", 3.4, "HIJ", 123456789)
    'retval = CalcStockMetrics(Worksheets("A"))

End Sub
