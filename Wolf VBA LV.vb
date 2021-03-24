Sub LoboWallstreet():

    For Each ws In Worksheets

        ' Columnas / Etiquetas
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ' Variables Iniciales 
        Dim Ticker_Name As String
        Dim Last_Row As Long
        Dim Total_Ticker_Volume As Double
        Total_Ticker_Volume = 0
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        Dim Yearly_Open As Double
        Dim Yearly_Close As Double
        Dim Yearly_Change As Double
        Dim Previous_Amount As Long
        Previous_Amount = 2
        Dim Percent_Change As Double
       

        ' Ultima Celda 
        Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To Last_Row

            ' Sumar al Ticker 
            Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
            ' Checar la Ubicación 
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


                ' Nombre del Ticker 
                Ticker_Name = ws.Cells(i, 1).Value
                ' Print Nombre Ticker en Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                ' Print Total del Ticker en Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                ' Resetear el Total del Ticker 
                Total_Ticker_Volume = 0

                ' Establecer Yearly Open, Yearly Close y Yearly Change Name
                Yearly_Open = ws.Range("C" & Previous_Amount)
                Yearly_Close = ws.Range("F" & i)
                Yearly_Change = Yearly_Close - Yearly_Open
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

                ' Determinar Percent Change
                If Yearly_Open = 0 Then
                    Percent_Change = 0
                Else
                    Yearly_Open = ws.Range("C" & Previous_Amount)
                    Percent_Change = Yearly_Change / Yearly_Open
                End If
                
                ' Formato doble para Incluir % Symbol And Two Decimal Places
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change

                ' Formato Condicional  Positivo(Verde) / Negativo (Rojo)
                If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
            
                ' Sumar uno al Summary Table Row
                Summary_Table_Row = Summary_Table_Row + 1
                Previous_Amount = i + 1
                End If
            Next i
         Next ws

End Sub