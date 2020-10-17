Attribute VB_Name = "Module1"
Sub VBA_allStreetMainMacro()
'Sheets(Sheets.Count).Select
    For Each ws In Worksheets
        ws.Select
        Call Resumen1
        ws.Columns("A:Q").AutoFit
    Next ws
MsgBox "Macro Succesed"
End Sub

Sub Resumen1()
'Declaration variables
Dim i, lRow, lLastRow As Long
Dim dOpen, dYearlyChange As Double
Dim sTickerGI, sTickerGD, sTicketGTV As String
Dim dVolumeGI, dVolumeGD As Double
Dim lVolumeGTV As LongLong
Dim lVolume As LongLong

'Headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Volume"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
'------------------------
lLastRow = Cells(Rows.Count, 1).End(xlUp).Row
lRow = 1

For i = 2 To lLastRow 'Main Loop
    If i = 2 Then 'First row step conditions
        dOpen = Cells(i, 3).Value
        lVolume = Cells(i, 7).Value
        
        lRow = lRow + 1
        sTickerGI = Cells(i, 1).Value
        sTickerGD = sTickerGI
        sTicketGTV = sTickerGD
        dVolumeGI = 0
        dVolumeGD = 0
        lVolumeGTV = 0
        Cells(lRow, 9).Value = Cells(i, 1).Value 'Ticker
    Else
        If i = lLastRow Then
            'Last row steps conditions - Ticker
            dYearlyChange = Cells(i, 6).Value - dOpen
            Cells(lRow, 10).Value = dYearlyChange
            If dYearlyChange > 0 Then 'Assign Interior colors
                Cells(lRow, 10).Interior.ColorIndex = 4 'GREEN
            Else
                Cells(lRow, 10).Interior.ColorIndex = 3 'RED
            End If
            
            'Yearly Charge % calculation
            Cells(lRow, 11).Value = dYearlyChange / dOpen
            Cells(lRow, 11).NumberFormat = "0.00%"
            
            'Volume acummulation last ticker
            lVolume = lVolume + Cells(i, 7).Value
            Cells(lRow, 12).Value = lVolume
            
            'Greatest  Increase % calculation
            If dVolumeGI < Cells(lRow, 11).Value Then
                dVolumeGI = Cells(lRow, 11).Value
                sTickerGI = Cells(lRow, 9).Value
            End If
            
            'Greates Decrese % calculation
            If dVolumeGD > Cells(lRow, 11).Value Then
                dVolumeGD = Cells(lRow, 11).Value
                sTickerGD = Cells(lRow, 9).Value
            End If
            
            'Gratest Total Volumen calculation
            If lVolumeGTV < Cells(lRow, 12).Value Then
                lVolumeGTV = Cells(lRow, 12).Value
                sTicketGTV = Cells(lRow, 9).Value
            End If
            
            
        Else
            'Rows different that 1st and last step conditions
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                dYearlyChange = Cells(i - 1, 6).Value - dOpen
                Cells(lRow, 10).Value = dYearlyChange
                If dYearlyChange > 0 Then ''Assign Interior colors
                    Cells(lRow, 10).Interior.ColorIndex = 4 'GREEN
                Else
                    Cells(lRow, 10).Interior.ColorIndex = 3 'RED
                End If
                
                'Yearly Charge % calculation
                If dOpen = 0 Then
                    Cells(lRow, 11).Value = dYearlyChange / 1
                Else
                    Cells(lRow, 11).Value = dYearlyChange / dOpen
                End If
                
                'Yearly Charge % calculation
                Cells(lRow, 11).NumberFormat = "0.00%"
                Cells(lRow, 12).Value = lVolume 'Total Stock Volume
                
                'Greates Volumen Increase
                If dVolumeGI < Cells(lRow, 11).Value Then
                    dVolumeGI = Cells(lRow, 11).Value
                    sTickerGI = Cells(lRow, 9).Value
                End If
                
                'Greates Decrese % calculation
                If dVolumeGD > Cells(lRow, 11).Value Then
                    dVolumeGD = Cells(lRow, 11).Value
                    sTickerGD = Cells(lRow, 9).Value
                End If
                
                'Gratest Total Volumen calculation
                If lVolumeGTV < Cells(lRow, 12).Value Then
                    lVolumeGTV = Cells(lRow, 12).Value
                    sTicketGTV = Cells(lRow, 9).Value
                End If
                
                'Ticker counter
                lRow = lRow + 1
                Cells(lRow, 9).Value = Cells(i, 1).Value 'Ticker
                
                dOpen = Cells(i, 3).Value 'dOpen
                lVolume = Cells(i, 7).Value
            Else
                'When ticker doesnt change volume accumulation continues
                lVolume = lVolume + Cells(i, 7).Value
            End If
        End If
    End If
Next
'Greatest calculation
Range("P2").Value = sTickerGI
Range("Q2").Value = dVolumeGI
Range("Q2").NumberFormat = "0.00%"
Range("P3").Value = sTickerGD
Range("Q3").Value = dVolumeGD
Range("Q3").NumberFormat = "0.00%"
Range("P4").Value = sTicketGTV
Range("Q4").Value = lVolumeGTV
End Sub

