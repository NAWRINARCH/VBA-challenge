
# VBA Script for Analyzing Quarterly Data

 This VBA script is designed to loop through all the stocks for each quarter and output the following information:
 1. The Ticker symbol
 2. Quarterly change from the opening price at the at the beginning of a given quarter to the closing price at the end of the quarter.
 3. the percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
 4. The total stock volume of the stock.

# Instructions:
1.Open the Excel workbook containing the stock data.
2.Navigate to the VBA editor.
3.Insert a new module to add  the following script.
4. Run the script to analyze the quarterly stock data.

Sub ProcessQuarterlyData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim ticker As String
    Dim openPrice As Double, closePrice As Double, volume As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double, percentChange As Double
    Dim tickerDict As Object
    Dim summaryRow As Long
    Dim tickerKeys As Variant

  For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Q*" Then
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

            ' Create dictionary to store data
            Set tickerDict = CreateObject("Scripting.Dictionary")

            ' Collect data from each row
            For i = 2 To lastRow
            If Not tickerDict.exists(ticker) Then
                    tickerDict.Add ticker, Array(openPrice, closePrice, volume)
                Else
                    Dim arr As Variant
                      End If
            Next i
    ' Add headers for the output columns in the current worksheet
            ws.Cells(1, 10).Value = "Ticker"
            ws.Cells(1, 11).Value = "Total Volume"
            ws.Cells(1, 12).Value = "Quarterly Change ($)"
            ws.Cells(1, 13).Value = "Percent Change (%)"
            ws.Cells(1, 15).Value = "Greatest % Increase"
            ws.Cells(1, 16).Value = "Greatest % Decrease"
            ws.Cells(1, 17).Value = "Greatest Total Volume"

            summaryRow = 2
            ' Get keys from the dictionary
            tickerKeys = tickerDict.Keys
     
     ' Write summary data to the current worksheet and calculate metrics
      For j = LBound(tickerKeys) To UBound(tickerKeys)
            ticker = tickerKeys(j)
            Dim data As Variant
            data = tickerDict(ticker)

            openPrice = data(0)
            closePrice = data(1)
            totalVolume = data(2)
            quarterlyChange = closePrice - openPrice

            

            ' Collect data from each row
    'Apply conditional formatting
    Sub ApplyConditionalFormatting(sheet As Worksheet, rangeStr As String, title As String)
    With sheet.Range(rangeStr)
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
        .FormatConditions(1).Interior.Color = RGB(144, 238, 144)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        .FormatConditions(2).Interior.Color = RGB(255, 99, 71)
    End With
End Sub 

The total stock volume of the stock result is shown in the following image:
![screenshot 1 VBA](https://github.com/NAWRINARCH/VBA-challenge/assets/170464172/b583ab63-b6dd-4a20-a635-c18a4a0aa06b)  


The stock with the "Greatest % increase", " Greatest percent decrease", and "Greatest total volume" is shown in the following image:
![screenshot 2 VBA](https://github.com/NAWRINARCH/VBA-challenge/assets/170464172/f44ef1bf-e6fd-4974-bfd5-fe6a60df4046)




