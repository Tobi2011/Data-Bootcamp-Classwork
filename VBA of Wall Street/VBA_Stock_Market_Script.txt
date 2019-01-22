Sub Stock_Market():
   'LOOP THROUGH EACH SHEET
    For Each ws In Worksheets
    'INITIALIZE ALL THE VARIABLES
        r = 1
        Tot_vol = 0
        Opn = 0
        Cls = 0
        Y_Change = 0
        P_Change = 0
        
        High_P = 0
        Low_P = 0
        High_vol = 0
    'CREATE HEADINGS AND LABELS FOR SUMMARY TABLE
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        ws.Cells(1, 16) = "Ticker"
        ws.Cells(1, 17) = "Value"
        ws.Cells(2, 15) = "Greatest % Increase"
        ws.Cells(3, 15) = "Greatest % Decrease"
        ws.Cells(4, 15) = "Greatest Total Volume"
        'LOOP THROUGH ALL THE ROWS IN SHEET
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            'DEFINE FIRST OPEN STOCK
            If i = 2 Then
                Opn = ws.Cells(i, 3)
            'IF THE THE NEXT STOCK DOES NOT MATCH THE CURRENT STOCK THEN PROCEED TO SUMMARY FOR CURRENT STOCK
            ElseIf ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
                Tot_vol = Tot_vol + ws.Cells(i, 7)
                'ADD STOCK VOLUME AND TICKER TO SUMMARY TABLE
                r = r + 1
                ws.Cells(r, 12) = Tot_vol
                ws.Cells(r, 9) = ws.Cells(i, 1)
                Cls = ws.Cells(i, 6)
                
               'CALCULATE AND DISPLAY YEARLY CHANGE
                Y_Change = Cls - Opn
                ws.Cells(r, 10) = Y_Change
                'CALCULATE AND DISPLAY PERCENTAGE CHANGE
                'THERE IS AN ZERO STOCK THAT MESSES WITH THE DATA AND REMOVE IT
                If Opn <> 0 Then
                    P_Change = Y_Change / Opn
                        If P_Change < 0 Then
                            ws.Cells(r, 11).Interior.Color = RGB(255, 0, 0)
                        ElseIf P_Change > 0 Then
                            ws.Cells(r, 11).Interior.Color = RGB(0, 255, 0)
                        End If
                Else:
                    P_Change = 0
                End If
                ws.Cells(r, 11) = FormatPercent(P_Change, 9)
                    'CALCULATE AND DISPLAY HIGHEST/LOWEST PERCENTAGE CHANGE
                    If P_Change > High_P Then
                        High_P = P_Change
                        ws.Cells(2, 17) = FormatPercent(High_P, 9)
                        ws.Cells(2, 16) = ws.Cells(i, 1)
                    
                    ElseIf P_Change < Low_P Then
                        Low_P = P_Change
                        ws.Cells(3, 17) = FormatPercent(Low_P, 9)
                        ws.Cells(3, 16) = ws.Cells(i, 1)
                    End If
                    'OBTAIN AND DISPLAY HIGHEST VOLUME
                    If Tot_vol > High_vol Then
                        High_vol = Tot_vol
                        ws.Cells(4, 17) = High_vol
                        ws.Cells(4, 16) = ws.Cells(i, 1)
                    End If
                    'PREPARE FOR NEXT STOCK
                    Tot_vol = 0
                    Opn = ws.Cells(i + 1, 3)
            'KEEP ADDING VOLUME IF THE CURRENT AND FUTURE STOCK MATCH
            Else
                Tot_vol = Tot_vol + ws.Cells(i, 7)
            End If
        Next i
        'MAKE HEADINGS FIT NICELY IN CELLS
        ws.Columns("A:T").AutoFit
    Next ws
End Sub