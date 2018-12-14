Attribute VB_Name = "Module1"
Sub HomeWork_Click()

'Declare Variables

Dim CurrentTicker As String
Dim NextTicker As String
Dim TotalStockVolume As LongLong
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim NextRow As Integer
Dim YearlyChange As Double
Dim PercentChange As Double
Dim Years As Variant
Dim Year As Variant

'Array to go through worksheets
Years = Array("2016", "2015", "2014")

'Loop through to move to the worksheet in the array
For Each Year In Years
ThisWorkbook.Worksheets(Year).Activate

'Set initial values to variables

    CurrentTicker = ""
    NextTicker = ""
    TotalStockVolume = 0
    NextRow = 2
    
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
    Cells(NextRow, 14).Value = OpenPrice
    Cells(NextRow, 15).Value = ClosePrice
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
High = Cells(2, 11).Value
Low = Cells(2, 11).Value

'Loop through the results from above loop code.

For j = 2 To LastResultRow

'Below if evaluate greatest increase by percentage
    If Cells(j, 11) > High Then
        High = Cells(j, 11).Value
        HighTicker = Cells(j, 10).Value
        Cells(2, 19).Value = HighTicker
        Cells(2, 20).Value = High
    End If

'Below if evaluate least decrease by percentage
    If Cells(j, 11) < Low Then
        Low = Cells(j, 11).Value
        LowTicker = Cells(j, 10).Value
        Cells(3, 19).Value = LowTicker
        Cells(3, 20).Value = Low
    End If
    
'Below if evaluate greatest volume
    If Cells(j, 13) > HighVolume Then
        HighVolume = Cells(j, 13).Value
        HighVolumeTicker = Cells(j, 10).Value
        Cells(4, 19).Value = HighVolumeTicker
        Cells(4, 20).Value = HighVolume
    End If


Next j


Next Year


End Sub




