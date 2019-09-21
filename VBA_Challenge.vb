Sub VBA_Challenge()

    'Enable Script to run on every Worksheet
    For Each ws In Worksheets

        'Summary Table Column Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        'Summary Table Row Labels
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        'Declare Variables and Define Values
        Dim Ticker_Name As String
        Dim Yearly_Open As Double
        Dim Yearly_Close As Double

        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0

        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        Dim Yearly_Open_Row As Long
        Yearly_Open_Row = 2
       
       'Determine Last Row of Data Table
        Dim Last_Row As Long
        Last_Row = ws.Cells(Rows.Count,1).End(xlUp).Row

        For i = 2 To Last_Row

            IF (ws.Cells(i + 1,1).Value <> ws.Cells(i,1).Value) Then

                'Set Ticker Name and Total Stock Volume Values
                Ticker_Name = ws.Cells(i,1).Value
                Total_Stock_Volume = Total_Stock_Volume + Cells(i,7).Value

                'Display Values in Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

                'Reset Total Stock Volume
                Total_Stock_Volume = 0

                'Set Yearly Open, Close and Change Values
                Yearly_Open = ws.Range("C" & Yearly_Open_Row).Value
                Yearly_Close = ws.Range("F" & i).Value
                Yearly_Change = Yearly_Close - Yearly_Open

                'Display Value in Summary Table
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

                'Set Percent Change Value
                IF (Yearly_Open = 0) Then

                    Percent_Change = 0

                Else

                    Yearly_Open = ws.Range("C" & Yearly_Open_Row).Value
                    Percent_Change = Yearly_Change / Yearly_Open

                End IF

                'Display Value in Summary Table with Proper Formatting
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

                'Conditional Formatting Showing Positive Values as Green and Negative Values as Red
                IF (ws.Range("J" & Summary_Table_Row).Value >= 0) Then

                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4

                Else

                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3

                End IF

                'Set New Summary Table Row and Yearly Open Row Values for Next Iteration
                Summary_Table_Row = Summary_Table_Row + 1
                Yearly_Open_Row = i + 1
            
            Else

                Total_Stock_Volume = Total_Stock_Volume + Cells(i,7).Value

            End IF

        Next i

        'Determine Last Row of Summary Table
        Dim Last_Row2 As Long
        Last_Row2 = ws.Cells(Rows.Count,9).End(xlUp).Row

        For i = 2 To Last_Row2

            'Determine Ticker with Greatest % Increase and Display Values in Summary Table
            IF (ws.Range("K" & i).Value > ws.Range("Q2").Value) Then
                
                ws.Range("Q2") = ws.Range("K" & i).Value
                ws.Range("P2").Value = ws.Range("I" & i).Value

            End IF

            'Determine Ticker with Greatest % Decrease and Display Values in Summary Table
            IF (ws.Range("K" & i).Value < ws.Range("Q3").Value) Then

                ws.Range("Q3").Value = ws.Range("K" & i).Value
                ws.Range("P3").Value = ws.Range("I" & i).Value

            End IF

            'Determine Ticker with Greatest Total Stock Volume and Display Values in Summary Table
            IF (ws.Range("L" & i).Value > ws.Range("Q4").Value) Then

                ws.Range("Q4").Value = ws.Range("L" & i).Value
                ws.Range("P4").Value = ws.Range("I" & i).Value

            End IF

        Next i

        'Format Cells Q2 and Q3 to Include % and Two Decimal Places
        ws.Range("Q2:Q3").NumberFormat = "0.00%"

        'Format Summary Table Columns to Auto Fit
        ws.Columns("I:Q").AutoFit

    Next ws

End Sub