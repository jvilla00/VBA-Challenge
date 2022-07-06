Attribute VB_Name = "Module1"

Sub multiyearstock()
For Each ws In Worksheets

    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim percentChange As Double
    Dim stockVal As Double
    Dim placeCounter As Integer

    openPrice = ws.Range("C2").Value
    placeCounter = 2
    stockVal = 0
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"



    For i = 2 To 22771
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            closePrice = ws.Cells(i, 6).Value
            ticker = ws.Cells(i, 1).Value
            ws.Cells(placeCounter, 10).Value = ticker
            ws.Cells(placeCounter, 11).Value = (closePrice - openPrice)
            ws.Cells(placeCounter, 12).Value = (closePrice - openPrice) / openPrice

            stockVal = stockVal + ws.Cells(i, 7).Value
            ws.Cells(placeCounter, 13).Value = stockVal
            placeCounter = placeCounter + 1
            stockVal = 0
            openPrice = ws.Cells(i + 1, 3).Value

        Else
            stockVal = stockVal + ws.Cells(i, 7).Value

        End If

    Next i
Next ws
End Sub



