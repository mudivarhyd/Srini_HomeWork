Attribute VB_Name = "Module1"
Sub HomeWork_Click()

'Declare Variables

Dim CurrentTicker As String
Dim NextTicker As String
Dim TotalStockVolume As Double
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim NextRow As Integer
Dim YearlyChange As Double
Dim PercentChange As Double
Dim Years As Variant
Dim Year As Variant
Dim WS_Count As Integer
Dim W As Integer



WS_Count = ActiveWorkbook.Worksheets.Count
'Loop through all of the worksheets in active workbook
For W = 1 To WS_Count
ThisWorkbook.Worksheets(W).Activate
'Set initial values to variables

    CurrentTicker = ""
    NextTicker = ""
    TotalStockVolume = 0
    NextRow = 2
    
'Set Header rows and Labels
Cells(1, 10) = "Ticker"
Cells(1, 11) = "Yearly Change"
Cells(1, 12) = "Percent Change"
Cells(1, 13) = "Total Stock Volume"
ThisWorkbook.Worksheets(W).Range("J1:M1").Font.Bold = True
ThisWorkbook.Worksheets(W).Range("J1:M1").Columns.AutoFit
Cells(1, 16) = "Ticker"
Cells(1, 17) = "Value"
ThisWorkbook.Worksheets(W).Range("P1:Q1").Font.Bold = True
ThisWorkbook.Worksheets(W).Range("P1:Q1").Columns.AutoFit
Cells(2, 15) = "Greatest % Increase"
Cells(3, 15) = "Greatest % Decrease"
Cells(4, 15) = "Greatest Total Value"
ThisWorkbook.Worksheets(W).Range("O2:O4").Font.Bold = True
ThisWorkbook.Worksheets(W).Range("O2:O4").Columns.AutoFit
    
'Assign total row count to the variable for Column A and initialize variables

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
CurrentTicker = Cells(2, 1).Value
OpenPrice = Cells(2, 3).Value
    
' Loop through all rows assigned to the lastrow variable. To Catch last sticker incrimented lastrow by 1

For i = 2 To lastrow + 1


    NextTicker = Cells(i, 1).Value
    
If CurrentTicker = NextTicker Then
    TotalStockVolume = Cells(i, 7) + TotalStockVolume
    NextTicker = Cells(i, 1).Value
    ClosePrice = Cells(i, 6).Value
'To Catch the value of last ticker incrimented lastrow by one hence valuating for NextTicker for Null value
ElseIf CurrentTicker <> NextTicker Or NextTicker = "" Then
    Cells(NextRow, 9).Value = CurrentYear
    Cells(NextRow, 10).Value = CurrentTicker
    YearlyChange = (ClosePrice - OpenPrice)
' To handle divisor by 0 error
        If OpenPrice = 0 Then
            PercentChange = 0
        Else
            PercentChange = (ClosePrice - OpenPrice) / OpenPrice
        End If
    Cells(NextRow, 11).Value = YearlyChange
    Cells(NextRow, 12).Value = PercentChange
        If YearlyChange <= 0 Then
            Cells(NextRow, 11).Interior.ColorIndex = 3
        Else
            Cells(NextRow, 11).Interior.ColorIndex = 4
        End If
    Cells(NextRow, 13).Value = TotalStockVolume
    TotalStockVolume = 0
    CurrentTicker = NextTicker
    CurrentYear = NextYear
    OpenPrice = Cells(i, 3).Value
    NextRow = NextRow + 1
End If


Next i

' Below code is to evaluate high, low Increase and decrease as well as max volume for the year.

Dim LastResultRow As Long
Dim High As Double
Dim HighTicker As String
Dim Low As Double
Dim LowTicker As String
Dim HighVolume As Double
Dim HighVolumeTicker As String


High = 0
HighTicker = ""
Low = 0
LowTicker = ""
HighVolume = 0
HighVolumeTicker = ""

LastResultRow = Cells(Rows.Count, 11).End(xlUp).Row
High = Cells(2, 12).Value
Low = Cells(2, 12).Value

'Loop through the results from above loop code.

For j = 2 To LastResultRow

'Below if evaluate greatest increase by percentage
    If Cells(j, 12) > High Then
        High = Cells(j, 12).Value
        HighTicker = Cells(j, 10).Value
        Cells(2, 16).Value = HighTicker
        Cells(2, 17).Value = High
    End If

'Below if evaluate least decrease by percentage
    If Cells(j, 12) < Low Then
        Low = Cells(j, 12).Value
        LowTicker = Cells(j, 10).Value
        Cells(3, 16).Value = LowTicker
        Cells(3, 17).Value = Low
    End If
    
'Below if evaluate greatest volume
    If Cells(j, 13) > HighVolume Then
        HighVolume = Cells(j, 13).Value
        HighVolumeTicker = Cells(j, 10).Value
        Cells(4, 16).Value = HighVolumeTicker
        Cells(4, 17).Value = HighVolume
    End If


Next j
ThisWorkbook.Worksheets(W).Range("Q2:Q4").Columns.AutoFit

Next W


End Sub




