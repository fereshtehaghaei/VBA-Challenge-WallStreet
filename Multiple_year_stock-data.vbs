Attribute VB_Name = "Module1"
Sub Stock_Market()

Dim WS As Worksheet

'Loop through all worksheets
For Each WS In ActiveWorkbook.Worksheets
WS.Activate

    'Setting Initial Variable new table Moderate Solution
    Dim Ticker_Name As String
    Dim Yearly_Change As Double
    Dim Percentage_Change As Double
    Dim Total_Stock_Volume As Variant
    Dim LastRow As Long
    Dim i As Long
    Yearly_Change = 0
    Percentage_Change = 0
    
    'Creating Headers/Titles for 1st Summary Table
    WS.Range("I1").Value = "Ticker"
    WS.Range("I1").Font.Bold = True
    
    WS.Range("J1").Value = "Yearly Change"
    WS.Range("J1").Font.Bold = True
    
    WS.Range("K1").Value = "Percentage Change"
    WS.Range("K1").Font.Bold = True
    
    WS.Range("L1").Value = "Total Stock Volume"
    WS.Range("L1").Font.Bold = True
    
    'Setting Variables for original table
    Dim open_price As Double
    Dim close_price As Double
    open_price = 0
    close_price = 0
    
    'Creating Header/Titles 2nd Summary Table for Hard solution part
    WS.Range("P1").Value = "Ticker"
    WS.Range("P1").Font.Bold = True
    
    WS.Range("Q1").Value = "Value"
    WS.Range("Q1").Font.Bold = True
    
    WS.Range("O2").Value = "Greatest % Increase"
    WS.Range("O2").Font.Bold = True
    
    WS.Range("O3").Value = "Greatest % Decrease"
    WS.Range("O3").Font.Bold = True
    
    WS.Range("O4").Value = "Greatest Total Volume"
    WS.Range("O4").Font.Bold = True
    
    ' Set new Ticker Variables
    Dim Max_Ticker_Name As String
    Dim Min_Ticker_Name As String
    Dim Max_Ticker_Volume As String
    Max_Ticker_Name = ""
    Min_Ticker_Name = ""
    Max_Ticker_Volume = ""
    
    ' Set new Value Variables
    Dim Max_Value As Double
    Dim Min_Value As Double
    Dim Max_Volume As Double
    Max_Value = 0
    Min_Value = 0
    Max_Volume = 0
    
    'Track of data in summary table
    Dim Summary_Table As Long
    Summary_Table = 2
    
    
    'setting initial value for open price
    open_price = WS.Cells(2, 3).Value
    
    'Finding the Last Row of each data table
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
    'looping from beginning of row2 till it reaches the lastrow
    For i = 2 To LastRow
        
    'check to see if we are still within the same ticker name if it is not...
    If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
    
    'Set the ticker name value
             Ticker_Name = WS.Cells(i, 1).Value
             
    'Print the ticker name in to summary table
            WS.Range("I" & Summary_Table).Value = Ticker_Name
    
    'Calculating Change in price value
             close_price = WS.Cells(i, 6).Value
             
             Yearly_Change = close_price - open_price
             
    'Print the Yearly Change
            WS.Range("J" & Summary_Table).Value = Yearly_Change
'-----------------------------------------------------------------------------------------------------------------
'Beginning conditional formatting section that will highlight positive change in green and negative change in red
'-----------------------------------------------------------------------------------------------------------------
                If Yearly_Change > 0 Then
                
                'fill the column cell value color with Green
                    WS.Range("J" & Summary_Table).Interior.ColorIndex = 4
                    
                ElseIf Yearly_Change <= 0 Then
                
                'fill the column cell value color with Red
                    WS.Range("J" & Summary_Table).Interior.ColorIndex = 3
                End If
'---------------------------------------------------------------------------------------------------------------
'End of conditional formatting section that will highlight positive change in green and negative change in red
'---------------------------------------------------------------------------------------------------------------
    
    'Set Percentage Chage value
                If open_price <> 0 Then
        
                    Percentage_Change = (Yearly_Change / open_price)
                   
                End If
                                                   
    'Print percentage change into summary table
                Range("K" & Summary_Table).Value = Percentage_Change
    
    'Formatting cell to have % sign since we have this NumberFormat formula we do not need to multiply our Percentage_Change with 100
                Range("K" & Summary_Table).NumberFormat = "0.00%"
            
     'Set stock volume value
                Total_Stock_Volume = Total_Stock_Volume + WS.Cells(i, 7).Value
                  
    'Print the Stock Volumne Total into summary table
                WS.Range("L" & Summary_Table).Value = Total_Stock_Volume
    
    
    'Add one to the Summary table so once it finds a different ticker number it sets values into next row
                Summary_Table = Summary_Table + 1
             
    '*Important*Captures the next Ticker's open price in order to add correctly
                open_price = WS.Cells(i + 1, 3).Value
    
    'Reset the Ticker name
                'Ticker_Name = " "
    
    'Reset values
                'Total_Stock_Volume = 0
                Yearly_Change = 0
                'close_price = 0
    
'---------------------------------------------------------------------------------------------------------------
'               2nd part of Solution "Hard_Solution"
'---------------------------------------------------------------------------------------------------------------
    
        ' <leave comment here>
        If (Percentage_Change > Max_Value) Then
                Max_Value = Percentage_Change
                Max_Ticker_Name = Ticker_Name
                
            
        ElseIf (Percentage_Change < Min_Value) Then
                Min_Value = Percentage_Change
                Min_Ticker_Name = Ticker_Name
    
        End If
            
    
        ' <leave comment here>
        If (Total_Stock_Volume > Max_Volume) Then
                Max_Volume = Total_Stock_Volume
                Max_Ticker_Volume = Ticker_Name
        
        End If
        
    'Reset value to zero for second summary table
        Total_Stock_Volume = 0
    
    'If the cells immediately following a row is the same ticker name do else....
       
    Else
                Total_Stock_Volume = Total_Stock_Volume + WS.Cells(i, 7).Value
                       
    End If
    
    
    Next i
           
            
    'Printing Greatest % increase in Value column
            WS.Range("Q2").Value = Max_Value
            WS.Range("Q2").NumberFormat = "0.00%"
            
    'Printing Greatest % Decrease in Value column
            WS.Range("Q3").Value = Min_Value
            WS.Range("Q3").NumberFormat = "0.00%"
            
    'Printing Greatest % increase in Ticker column
            WS.Range("P2").Value = Max_Ticker_Name
            
    'Printing Greatest % Decrease in Ticker column
            WS.Range("P3").Value = Min_Ticker_Name
            
    'Printing Greatest Total Volume in Value column
            WS.Range("Q4").Value = Max_Volume
            
    'Printing Greatest Total Volume in Value column
            WS.Range("P4").Value = Max_Ticker_Volume
Next WS

End Sub
