Sub StockTickerAnalysis()

'*******************************************************************************
'*******************************************************************************
'*******************************************************************************

'*******************************************************************************
'**************** VARIABLE DECLARATION *************************************
'*******************************************************************************

Dim rowIndex As Integer
Dim year As Integer
Dim totalVolume As Integer

Dim openPrice As Double
Dim closePrice As Double
Dim percentChange As Double

Dim Symbol As String

'*******************************************************************************
'******************* VARIABLE INITIATION *************************************
'*******************************************************************************
year = Int(Left(Cells(2, 2).Value, 4))
openPrice = 0#
closePrice = 0#
Symbol = Cells(2, 1).Value

'*******************************************************************************
'**************** FOR LOOP FOR ANALYSIS ************************************
'*******************************************************************************

For rowIndex = 2 To lastRow

'*******************************************************************************
'*Checks to see if the value of the Year Changes.  If there is a change,
'a flag is thrown and documented on the active spreadsheet
'*******************************************************************************
    If Int(Left(Cells(rowIndex, 2).Value, 4)) <> year Then
        MsgBox ("More than One Year in this table")
        Cells(3, 10).Value = "More than one Ticker Symbol"
    End If

'*******************************************************************************
'*Checks to see if the value of the Ticker Symbol Changes.  If there is a change,
'a flag is thrown and documented on the active spreadsheet
'*******************************************************************************

    If Cells(rowIndex, 1).Value <> Symbol Then
        MsgBox ("More than One Symbol")
        Cells(3, 11).Value = "More than one Ticker Symbol"
    End If


'*******************************************************************************
'*Checks to see if the value of the Ticker Symbol Changes.  If there is a change,
'a flag is thrown and documented on the active spreadsheet
'*******************************************************************************
    'Find OpenPrice
    
    If rowIndex <= 10 Then
        MsgBox (Int(Cells(rowIndex, 2).Value))
    End If
    
    If Int(Right(Cells(rowIndex, 2).Value, 3)) = 101 Then
        openPrice = Cells(rowIndex, 3).Value
        MsgBox ("Open Price: " + Str(openPrice))
    End If



Next rowIndex

'*******************************************************************************
'**************** Assign Results to Cells on Active Sheet **********************
'*******************************************************************************

Cells(2, 10).Value = year
Cells(2, 11).Value = Symbol



'*******************************************************************************
'*******************************************************************************
'*******************************************************************************

End Sub

