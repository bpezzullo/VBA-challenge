Attribute VB_Name = "Module1"
Global Ticker_col, Ticker_open_col, Ticker_close_col, Ticker_volume_col As Integer
Global results_ticker_col, _
       results_chg_col, _
       results_per_col, _
       results_vol_col As Integer
' Global Ticker_array() As Long                               ' Setup an array to store working variables of the row count
                                                              ' I opted to not go in this direction because I wasn't certain on
                                                              ' processing time vs using open cells in the spreadsheet.
       

Sub Main()

Dim i As Integer
Dim sCellVal As String
Dim R As Range

For Each R In Range("A1:G1")                    ' Set the global variables based on the titles of the column

    sCellVal = R.Text
    
    If sCellVal Like "<ticker>" Then Ticker_col = R.Column
    
    If sCellVal Like "<open>" Then Ticker_open_col = R.Column

    If sCellVal Like "<close>" Then Ticker_close_col = R.Column

    If sCellVal Like "<vol>" Then Ticker_volume_col = R.Column
    
    Next R

' Set the global variables for the output.
results_ticker_col = 9
results_chg_col = 10
results_per_col = 11
results_vol_col = 12



' Loop through all the sheets  For i = 1 To 1ast sheet

For i = 1 To Application.Sheets.Count
    Worksheets(i).Activate
    Call Calculate

Next i


End Sub


Sub Calculate()

Dim Ticker_cell As Integer          'Ticker_cell keeps track on what row to store the results on the spreadsheet.
Dim start As String                 'Start is the current ticker name
Dim Found As String                 'Found is the next Ticker symbol found from the function
Dim rowct As Long                   'Rowct is used to determine where in the spreadsheet the Ticker changes.
Dim continue As Boolean             'Continue is used to control the loop
Dim Year As Integer                 ' The year to search for
Dim Volcell2 As String              ' Not needed
Dim Test As String
Dim cur_range As String
Dim temp As Double

'Initialize variables

Ticker_cell = 2                         ' Start with the second row
rowct = 1                               ' Start with 1
Found = Cells(1, rowct).Value           ' Set Found with the dummy label. I used the first cell
start = Found                           '
continue = True                         ' set the while flag to the initial value


' Inititialize spreadsheet with the labels and formating by calling Init_spread
Call Init_spread

' Loop through the ticker names in the first column (Ticker_col) until the end of the spreadsheet
' Call the function Identity_ticker to find the next ticker name.  Store the row number to determine where
' the comapny starts  in the spreadsheet for future subroutine calls
' Note: Data requirement is that the spreadsheet has been sorted by ticker and then date with each spreadsheet
' being for the same year.

Do While continue = True

 ' Need to find the row of the next Ticker symbol.   Call the Function Identify_ticker passing in the current sticker
 ' (start) and the rowcount (rowct) to start looking from.  The function returns the row count of the next ticker

 rowct = Identify_ticker(start, rowct)                   ' start from the last spot rowct
 
 Found = Cells(rowct, Ticker_col).Value                  ' Change of ticker symbol has been identified. Set the Found field
 Cells(Ticker_cell, results_ticker_col).Value = Found    ' Place in the spreadsheet at the appropriate location
                                                         '
 
 ' ReDim Preserve Ticker_array(Ticker_cell - 2) As Long  ' opted not to use an array
 ' Ticker_array(Ticker_cell - 2) = rowct                 ' Save the row in an array for later access
 Cells(Ticker_cell, 20).Value = rowct                    ' Save the row in an empty portion of the spreadsheet
 
 ' Increase the row for the next ticker symbol
 Ticker_cell = Ticker_cell + 1
 
 ' reset the Start variable for the while loop
 
 start = Found
 
 ' Are we at the end of the spreadsheet. If so jump out of the while loop
  'If Found = "AAN" Then continue = False               ' This is for debugging
  
  If Found = "" Then continue = False                   ' The last cell will return "".  If at end set the continue
                                                        ' flag to stop the while loop

 ' Get the next sticker
 Loop

 
 ' We have all the Tickers and where they start in the spreadsheet.  Time to analyze the calculated data to summarize results
 ' Loop through the tickers and add the Change, percent change, and total volume. Ticker_cell represents the end of the new
 ' ticker matrix.  I could of combined this with the above loop, but separated for a cleaner distinction of what each loop
 ' is doing.
 
 Call Cal_Ticker(Ticker_cell - 2)                         ' Calculate the change for the year, percent change, and total volume
                                                        ' Pass in the number of Tickers found in the sheet (Tciker_cell -
                                                        ' which is the row number -2 as we start at the second row)
 
 temp = get_max(11, 2, Ticker_cell)                     ' Find the maximum percent increase and return the Row number
                                                        ' from the new function get_max.  Pass in the row being examined
 
  Range("O2").Value = Cells(temp, results_ticker_col).Value          ' store the name of the ticker
  Range("P2").Value = Cells(temp, results_per_col).Value             ' store the percentage increase
 
 
 temp = get_min(11, 2, Ticker_cell)                     ' Find the worst percent increase (aka decrease) and return the Row
                                                        ' number from the new function get_min
 
  Range("O3").Value = Cells(temp, results_ticker_col).Value          ' store the name of the ticker
  Range("P3").Value = Cells(temp, results_per_col).Value             ' store the percentage decrease

' Find the maximum volume

 temp = get_max(12, 2, Ticker_cell)                     ' Find the maximum volume and return the Row number
                                                        ' from the new function get_max
 
  Range("O4").Value = Cells(temp, results_ticker_col).Value          ' store the name of the ticker
  Range("P4").Value = Cells(temp, results_vol_col).Value             ' store the Max volume

 ' Highlight the percentages and clean-up
 
 Call highlight(2, Ticker_cell)

End Sub

Function Identify_ticker(start As String, rowct As Long) As Long

' Find the next ticker

Dim Hold As String                                  ' Hold is a temporary field
Dim i As Long                                       ' variable to help count
Dim ticker_continue As Boolean                      ' Continue flag for while loop

i = rowct                                           ' Start at the last know location in the spreadsheet
ticker_continue = True                              ' Loop while continueis True

' Loop until the next sysmbol
Do While ticker_continue = True

 Hold = Cells(i, Ticker_col).Value
 If start <> Hold Then                              ' Found a new ticker symbol
    Identify_ticker = i                             ' return the row of the new ticker symbol
    ticker_continue = False                         ' Exit from the function
 Else
    i = i + 1                                       'step down to the next row
 End If

Loop

End Function

Sub Cal_Ticker(Ticker_cnter As Integer)
' this subroutine takes in the number of unique tickers and parses the summarized table

Dim ticker_start As Long                           ' Start of the data for ticker
Dim ticker_end, rowcnt As Long                     ' End of the data for a ticker.
Dim i As Integer
Dim result As Double


' Initialize rowcnt to the first row which is 2
 rowcnt = 2
 
 ' Loop from the start of the summarize table 2 to the last row (Ticker_cnter).
 For i = 2 To Ticker_cnter
 
  ' Find the maximum and minimum % increases plus volume for each ticker.
  ' In some cases the data doesn't have data at the beginning days and / or end  before the end of the year.
  ' In these cases the data will have 0 at the start or end of the data for each ticker.  Need to analyze the data to see if that is
  ' the case
 
  ' For each ticker find start of data area.  Eliminate starting and ending data with 0's for opening and closing.
  ' Also calculate the volume during this loop

  ' Initialize result to 0
    result = 0
    ticker_start = rowcnt

    ' loop through all the same tickers.

    Do While Cells(i, results_ticker_col).Value = Cells(rowcnt, Ticker_col).Value
    
        ' Skip the rows that are 0 open at the top of the ticker
        If (Cells(rowcnt, Ticker_open_col).Value = 0 And ticker_start = rowcnt) Then ticker_start = rowcnt + 1
        
        ' Skip the rows with 0 close at the end. Basically don't increase the ticker_end variable if close is a 0.
        If Cells(rowcnt, Ticker_close_col).Value <> 0 Then ticker_end = rowcnt     ' End of the ticker, skip the 0's at the end
        
        result = result + Cells(rowcnt, Ticker_volume_col).Value                    ' While we are looping through the ticker data sum up the volumes
        rowcnt = rowcnt + 1
    
        Loop
        
    If ticker_start < ticker_end Then                                   ' if the ticker_end is less than the ticker_start
                                                                        ' then the ticker has all 0's
    
        ' Take the difference between the ending price and the starting price
        Cells(i, results_chg_col).Value = Cells(ticker_end, Ticker_close_col).Value - Cells(ticker_start, Ticker_open_col).Value
        
        ' Take the percent change
        Cells(i, results_per_col).Value = Cells(i, results_chg_col).Value / Cells(ticker_start, Ticker_open_col).Value
    Else
        ' no data for the ticker has been stored.  Place 0's in the fields
        Cells(i, results_chg_col).Value = 0
        Cells(i, results_per_col).Value = 0

    End If
 
    Cells(i, results_vol_col).Value = result
    
 Next i
    

 

End Sub



Private Sub highlight(min As Integer, max As Integer)

' Highlight the cells with red or green bases on positive or negative gain.  For 0 leave with no coloring

Dim rng As Range, cell As Range
Dim rangeinfo As String


rangeinfo = cellval("J", min) + ":" + cellval("J", max)

Set rng = Range(rangeinfo)

' set all cells to no backfill
rng.Interior.ColorIndex = 0

For Each cell In rng

    If cell.Value > 0 Then
        cell.Interior.ColorIndex = 4
    End If
    If cell.Value < 0 Then
        cell.Interior.ColorIndex = 3
    End If
        
Next cell

rangeinfo = cellval("K", min) + ":" + cellval("K", max)

Set rng = Range(rangeinfo)

rng.NumberFormat = "0.00%"

Set rng = Range("T:T")
rng.ClearContents

End Sub


Function cellval(letter As String, row As Integer) As String

' Function to concatenate the Column and row

cellval = letter + CStr(row)

End Function


Function get_max(col As Integer, first_row As Integer, last_row As Integer) As Integer

' find the maximum value and return the row that the value was found in.

Dim i As Integer
Dim temp As Double

temp = Cells(first_row, col).Value              ' start with the first cell as the temp
get_max = first_row

For i = first_row To last_row                   ' loop through the column provided

If Cells(i, col).Value > temp Then              ' If value is greater than temp
    get_max = i                                 ' capture the row
    temp = Cells(i, col).Value                  ' capture the new temp value
    End If
Next i

End Function


Function get_min(col As Integer, first_row As Integer, last_row As Integer) As Integer
' find the minimum value and return the row that the value was found in.

Dim i As Integer
Dim temp As Double

temp = Cells(first_row, col).Value              ' start with the first cell as the temp
get_min = first_row

For i = first_row To last_row                   ' loop through the column provided

If Cells(i, col).Value < temp Then              ' If value is lesser than temp
    get_min = i                                 ' capture the row
    temp = Cells(i, col).Value                  ' capture the new temp value
    End If
Next i

End Function

Sub Init_spread()
' Clear where the results will be

Range("I:P").ClearContents

' Label and Format the result Columns

Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("L1") = "Volume"
Range("O1") = "Ticker"
Range("P1") = "Value"
Range("I1:P1").Font.Bold = True
Range("N2") = "Greatest % Increase"
Range("N3") = "Greatest % Decrease"
Range("N4") = "Greatest Volume"
Range("N2").Font.Bold = True
Range("N3").Font.Bold = True
Range("N4").Font.Bold = True
Range("N:N").ShrinkToFit = True
Range("J1:K1").WrapText = True
Range("J1:L1").ShrinkToFit = True
Range("P2:P3").NumberFormat = "0.00%"

End Sub






