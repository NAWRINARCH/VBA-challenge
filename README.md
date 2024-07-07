
# VBA Script for Analyzing Quarterly Data

 This VBA script is designed to loop through all the stocks for each quarter and output the following information:
 1. The Ticker symbol
 2.Quarterly change from the opening price at the at the beginning of a given quarter to the closing price at the end of the quarter.
 3.  the percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
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
