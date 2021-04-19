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
Dim LastRow As Long

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

LastRow = Cells(Rows.Count, "A").End(xlUp).Row

For rowIndex = 2 To LastRow

'*******************************************************************************
'Checks to see if the value of the Year Changes. If there is a change,
'a flag isthrown and documented on the active spreadsheet
'******************************************************************************
If Int(Left(Cells(rowIndex, 2).Value, 4)) <> year Then
Cells(3, 10).Value = "More than one Year"
Else: Cells(3, 10).Value = year
End If

'*******************************************************************************
'Checks to see if the value of the Ticker Symbol Changes. If there is a change,
'a flag is thrown and documented on the active spreadsheet
'******************************************************************************

If Cells(rowIndex, 1).Value <> Symbol Then
    Cells(3, 11).Value = "More than one Ticker Symbol"
End If
'*******************************************************************************
'Checks to see if the value of the Ticker Symbol Changes. If there is a change,
'a flag is thrown and documented on the active spreadsheet
'******************************************************************************
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




Sub Solution()

    Dim LastRow As Long
    Dim Counter As Long
    Dim Counter_Result As Long
    
    Dim Ticker As String
    
    Dim TotalStockVolume As Double
    Dim Open_Price As Double
    Dim Close_Price As Double
    
    Set WB = ActiveWorkbook
    'Set WS = WB.ActiveSheet
    Set WS = WB.Sheets
    
    
    For Each WS In ThisWorkbook.Worksheets

    
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
    Counter = 2
    
    Ticker = WS.Cells(Counter, 1)
    Open_Price = WS.Cells(Counter, 3)
    
    WS.Range("I2:L1000000").ClearContents
    WS.Range("I2:L1000000").ClearFormats
    
    Counter_Result = 2
    
    Do While Counter <= LastRow
        If WS.Cells(Counter, 1) <> Ticker Then
            
            WS.Cells(Counter_Result, 9) = Ticker
            WS.Cells(Counter_Result, 12) = TotalStockVolume
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
        
            Ticker = WS.Cells(Counter, 1)
            TotalStockVolume = WS.Cells(Counter, 7)
            Open_Price = WS.Cells(Counter, 3)
            Counter_Result = Counter_Result + 1
        Else
            TotalStockVolume = TotalStockVolume + WS.Cells(Counter, 7)
        End If
        
        Counter = Counter + 1
        
    Loop
    
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
