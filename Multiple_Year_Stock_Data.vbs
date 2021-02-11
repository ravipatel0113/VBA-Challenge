Attribute VB_Name = "Module1"
Sub Run()
    ' For running the same loop through all the shhets in the worksheets
    Dim ws, sht As Worksheet
    Application.ScreenUpdating = False 'For screenupdating before the loop
    For Each ws In Worksheets
        ws.Select
        Call main ' Main macro of the script
        Call Bonus_Part 'Bonus part of the assignment
    Next ws
    For Each sht In ThisWorkbook.Worksheets
        sht.Cells.EntireColumn.AutoFit 'Autofit the columns for better view
    Next sht
    Application.ScreenUpdating = True 'For screenupdating after the loop
    
End Sub

Sub main() 'Main part of the macro script

    
    Dim WS_Count As Double 'to get the last row
    Dim Ticker As String
    Dim i, j As Double
    j = 0  ' to get it back to the top
    Dim Volume As Double
    Volume = 0
    ' Table header format
    Range("J1").Value = "Ticker_Symbol"
    Range("K1").Value = "Yearly_Change"
    Range("L1").Value = "Percentage_Change"
    Range("M1").Value = "Total_Stock_Volume"
    Dim Yearly_change, Percentage_change As Double
    Yearly_change = 0
    Percentage_change = 0
                                 
    ' Keep track of the location for each ticker in the summary table
    Dim ST As Double 'Summary Table
    ST = 2
    WS_Count = Cells(Rows.Count, 1).End(xlUp).Row
    
    
    ' Loop through all tickers
    For i = 2 To WS_Count
        ' Check if we are still within the same ticker, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            Yearly_change = Cells(i, 6).Value - Cells(i - j, 3).Value ' Yearly change which is close-open
            
            If (Yearly_change = 0) Or (Cells(i - j, 3) = 0) Then
                Percentage_change = 0 'for the error of divisible by 0
            Else
                Percentage_change = Yearly_change / Cells(i - j, 3).Value 'if it is not 0
            End If
            
            ' Set the ticker name
            Ticker = Cells(i, 1).Value

            ' Add to the volume
            Volume = Volume + Cells(i, 7).Value 'add up the volume of the ticker

            ' Print the ticker in the Summary Table
            Range("J" & ST).Value = Ticker 'Display the ticker name

            ' Print the volume to the Summary Table
            Range("M" & ST).Value = Volume 'Display the total volume
           
      
            ' Reset the volume Total
            Volume = 0
            j = 0
            Cells(ST, 11).Value = Yearly_change 'Display the yearly change
            Cells(ST, 12).Value = Percentage_change 'Display the percentage change
            Cells(ST, 12).NumberFormat = "0.00%" 'Format the percentage value to percentage
           
           
            If Cells(ST, 11) >= 0 Then
                Cells(ST, 11).Interior.Color = RGB(0, 255, 0) 'Color formatting if the value is positive
            Else
                Cells(ST, 11).Interior.Color = RGB(255, 0, 0) 'Color formatting if the value is negative
            End If
            Cells(ST, 11).NumberFormat = "$#,##0.00" 'Format the value to currency
            If Cells(ST, 12) >= 0 Then
                Cells(ST, 12).Interior.Color = RGB(0, 255, 0) 'Color formatting if the value is positive
            Else
                Cells(ST, 12).Interior.Color = RGB(255, 0, 0) 'Color formatting if the value is negative
            End If
            
            
            
        ' If the cell immediately following a row is the same ticker...
         ' Add one to the summary table row
         ST = ST + 1
        Else

            ' Add to the Volume Total
             j = j + 1
            Volume = Volume + Cells(i, 7).Value
        
        End If

        Next i
      
   ' Next ws
   
   
End Sub

Sub Bonus_Part() 'Bonus part of the assignment

    Dim i As Double
    WS_Count = Cells(Rows.Count, 1).End(xlUp).Row 'get to the last row
    Dim G_P As Double 'Greatest percentage increase
    Dim G_T As String 'Greatest percentage increase ticker
    G_P = 0
    Range("O2").Value = "Greatest % Increase"
    Dim L_P As Double 'greatest percentage decrease
    Dim L_T As String 'Greatest percentage decrease ticker
    L_P = 0
    Range("O3").Value = "Greatesh % Decrease"
    Dim G_V As Double 'Greatest Total Volume
    Dim G_V_T As String 'Greatest Total Volume Ticker
    G_V = 0
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker Symbol"
    Range("Q1").Value = "Value"
    
    For i = 2 To WS_Count
    
        If Cells(i, 12).Value > G_P Then 'Find ticker with Greatest Increase in percentage
            G_P = Cells(i, 12).Value
            G_T = Cells(i, 10).Value
            Range("P2") = G_T ' Display the ticker symbol
            Range("Q2") = G_P ' Display the percentage value
            Range("Q2").NumberFormat = "0.00%" 'Formatting the Value to percentage
        End If
        
         If Cells(i, 12).Value < L_P Then 'Find ticker with Greatest decrease in percentage
            L_P = Cells(i, 12).Value
            L_T = Cells(i, 10).Value
            Range("P3") = L_T 'Display the ticker Symbol
            Range("Q3") = L_P 'Display the least percentage value
            Range("Q3").NumberFormat = "0.00%"
        End If
        
         If Cells(i, 13).Value > G_V Then 'Find the ticker with greatest total volume
            G_V = Cells(i, 13).Value
            G_V_T = Cells(i, 10).Value
            Range("P4") = G_V_T 'Display the ticekr symbol
            Range("Q4") = G_V ' Display the total volume of the ticker
        End If
Next i
    
End Sub
