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

HOMEWORK ASSIGNMENT SHOULD BE COMPLETE.

THE ONLY FILE IN THIS ASSIGNMENT SHOULD BE THIS README.MD FILE THAT CONTAINS
THE STATUS OF THE HOMEWORK AND THE RAW SCRIPT FILE.

THIS CODE HAS BEEN TESTED ON THE "ALPHABET" XLS BUT NOT THE MAIN FILE.

********************************************************************************
***************************** Current VBA Code *********************************
********************************************************************************


Sub StockTickerAnalysis()

'*******************************************************************************
'********************* VARIABLE DECLARATION ********************************
'*******************************************************************************
    Dim LastRow As Long
    Dim Counter As Long
    Dim Counter_Result As Long
    
    Dim Ticker As String
    
    Dim TotalStockVolume As Double
    Dim Open_Price As Double
    Dim Close_Price As Double
    
    Set WB = ActiveWorkbook
    Set WS = WB.Sheets
    
    
'*******************************************************************************
'******************FOR LOOP THAT TRAVERSES EACH WORKSHEET *************
'*******************************************************************************
    
    For Each WS In ThisWorkbook.Worksheets
    
    
'*******************************************************************************
'********** INITIALIZE EACH VARIABLE FOR PER WORKSHEET ANALYSIS *********
'*******************************************************************************

        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        Counter = 2
        Ticker = WS.Cells(Counter, 1)
        Open_Price = WS.Cells(Counter, 3)
        
        WS.Range("I2:L1000000").ClearContents
        WS.Range("I2:L1000000").ClearFormats
        
        Counter_Result = 2

'*******************************************************************************
'*********** TRAVERSE EACH ROW IN THE CURRENT ACTIVE SHEET *************
'*******************************************************************************
        Do While Counter <= LastRow
        
            'If the ticker in the current row is different than that of the previous
            'row, then set the Ticker variable to the new ticker and run through the
            ' Open / Close / total ticker and % Change algorithm.
            If WS.Cells(Counter, 1) <> Ticker Then
                WS.Cells(Counter_Result, 9) = Ticker
                WS.Cells(Counter_Result, 12) = TotalStockVolume
                Close_Price = WS.Cells(Counter - 1, 6)
                WS.Cells(Counter_Result, 10) = Close_Price - Open_Price
                
                'Based on the instructions, this section sets the colors depicting the
                'change in values
                
                If WS.Cells(Counter_Result, 10) < 0 Then
                    WS.Cells(Counter_Result, 10).Interior.Color = vbRed
                Else
                    WS.Cells(Counter_Result, 10).Interior.Color = vbGreen
                End If
                
                'Evaluates % change per ticker and formats the cell to show %age
                If Close_Price = 0 Then
                    If Open_Price = 0 Then
                        WS.Cells(Counter_Result, 11) = FormatPercent(0, 2)
                    Else
                        WS.Cells(Counter_Result, 11) = FormatPercent(-1, 2)
                    End If
                Else
                    WS.Cells(Counter_Result, 11) = FormatPercent((Close_Price - Open_Price) / Close_Price, 2)
                End If
            
                Ticker = WS.Cells(Counter, 1)
                TotalStockVolume = WS.Cells(Counter, 7)
                Open_Price = WS.Cells(Counter, 3)
                Counter_Result = Counter_Result + 1
            Else
                'This means that the the ticker symbol has not changed yet from row to row.
                'Therefore Ticker Sum can be updated here.
                TotalStockVolume = TotalStockVolume + WS.Cells(Counter, 7)
            End If
            
            Counter = Counter + 1
            
        Loop
        
        'By Closing the While Loop Above, we have figured out all the information required
        'by the current active worksheet.  This phase of the algorithm below displays the
        'results  of the algorithm above by assigning the results to empty cells in each
        'worksheet.
        
        Close_Price = WS.Cells(Counter - 1, 6)
        WS.Cells(Counter_Result, 10) = Close_Price - Open_Price
        
        If WS.Cells(Counter_Result, 10) < 0 Then
            WS.Cells(Counter_Result, 10).Interior.Color = vbRed
        Else
            WS.Cells(Counter_Result, 10).Interior.Color = vbGreen
        End If
        
        If Close_Price = 0 Then
            If Open_Price = 0 Then
                WS.Cells(Counter_Result, 11) = FormatPercent(0, 2)
            Else
                WS.Cells(Counter_Result, 11) = FormatPercent(-1, 2)
            End If
        Else
            WS.Cells(Counter_Result, 11) = FormatPercent((Close_Price - Open_Price) / Close_Price, 2)
        End If
        
        
        WS.Cells(Counter_Result, 9) = WS.Cells(Counter - 1, 1)
        WS.Cells(Counter_Result, 12) = TotalStockVolume
        
    Next WS

End Sub



End Sub


