********************************************************************************
****************************** Author Details **********************************
********************************************************************************

Author:  	Daneshmand, Ali
Course:  	Rutgers Data Science Boot Camp
Instructor:	Harneet
T.A.:		Gretel

Assignment Deadline:  April 17, 2021

********************************************************************************
****************************** Instructions ************************************
********************************************************************************

* Create a script that will loop through all the stocks for one year and output 
the following information.

  * The ticker symbol.
  * Yearly change from opening price at the beginning of a given year to the 
    closing price at the end of that year.
  * The percent change from opening price at the beginning of a given year to 
    the closing price at the end of that year.
  * The total stock volume of the stock.

* You should also have conditional formatting that will highlight positive change 
in green and negative change in red.


********************************************************************************
***************************** Homework Status **********************************
********************************************************************************
1.  Found A Ticker -- Need to Triage for more than 1 ticker and apply the 
    existing VBA script per ticker
2.  Found Year
3.  Issue finding "Open Price" and "Close Price" for one Ticker.  Need to apply 
    logic for each ticker
4.  Found Total Stock Volume for one Ticker.  Need to apply to more than one 
    ticker.
5.  Issue applying Code to all worksheets in one shot. Gretel reminded me that 
    there's a "Loop" command I can impleent that applies the same logic to ALL sheets.


********************************************************************************
***************************** Current VBA Code *********************************
********************************************************************************




Sub StockTickerAnalysis()

'*******************************************************************************
'*******************************************************************************
'*******************************************************************************

'*******************************************************************************
'**************** VARIABLE DECLARATION *************************************
'*******************************************************************************

Dim rowIndex As Long
Dim year As Integer
Dim totalVolume As LongLong
Dim lastRow As Long

Dim openPrice As Long
Dim closePrice As Long
Dim percentChange As Double
Dim yearChange As Double

Dim Symbol As String

'*******************************************************************************
'******************* VARIABLE INITIATION *************************************
'*******************************************************************************
year = Int(Left(Cells(2, 2).Value, 4))
openPrice = 0#
closePrice = 0#
Symbol = Cells(2, 1).Value
totalVolume = 0

'*******************************************************************************
'**************** FOR LOOP FOR ANALYSIS ************************************
'*******************************************************************************

lastRow = Cells(Rows.Count, "A").End(xlUp).Row

For rowIndex = 2 To lastRow


'*******************************************************************************
'*Checks to see if the value of the Year Changes.  If there is a change,
'a flag is thrown and documented on the active spreadsheet
'*******************************************************************************
    If Int(Left(Cells(rowIndex, 2).Value, 4)) <> year Then
        Cells(3, 10).Value = "More than one Year"
    Else
        Cells(3, 10).Value = year
    End If

'*******************************************************************************
'*Checks to see if the value of the Ticker Symbol Changes.  If there is a change,
'a flag is thrown and documented on the active spreadsheet
'*******************************************************************************

    If Cells(rowIndex, 1).Value <> Symbol Then
        Cells(3, 11).Value = "More than one Ticker Symbol"
    End If
    
'*******************************************************************************
'*Checks to see if the value of the Ticker Symbol Changes.  If there is a change,
'a flag is thrown and documented on the active spreadsheet
'*******************************************************************************
    'Find OpenPrice
    
    If Int(Right(Cells(rowIndex, 2).Value, 3)) = 101 Then
        openPrice = Cells(rowIndex, 3).Value
    End If

'*******************************************************************************
'Total Volume Calculation
'*******************************************************************************

    totalVolume = totalVolume + Cells(rowIndex, 7).Value

Next rowIndex

'*******************************************************************************
'**************** Assign Results to Cells on Active Sheet **********************
'*******************************************************************************

Cells(2, 10).Value = year
Cells(2, 11).Value = Symbol
Cells(2, 12).Value = yearChange
Cells(2, 13).Value = percentChange
Cells(2, 14).Value = totalVolume



'*******************************************************************************
'*******************************************************************************
'*******************************************************************************

End Sub


